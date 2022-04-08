VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form supplier_entry_form 
   BackColor       =   &H80000004&
   Caption         =   "Supplier Entry Form"
   ClientHeight    =   10935
   ClientLeft      =   10470
   ClientTop       =   465
   ClientWidth     =   10635
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   10635
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   7455
      Left            =   2760
      TabIndex        =   22
      Top             =   600
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   13150
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "SUPPLIER"
      TabPicture(0)   =   "supplier_entry_form.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(2)=   "Image1"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "PRODUCT"
      TabPicture(1)   =   "supplier_entry_form.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(2)=   "Image2"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "VIEW"
      TabPicture(2)   =   "supplier_entry_form.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label9"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label25"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label26"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label27"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label29"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label30"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Line2"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "ListView2"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "ListView1"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).ControlCount=   9
      Begin MSComctlLib.ListView ListView1 
         Height          =   2655
         Left            =   240
         TabIndex        =   65
         Top             =   960
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   4683
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "NAME"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "COMP. NAME"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "PHONE"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "ADDRESS"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "GSTIN"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "BANK"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "IFSC CODE"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "ACCOUNT NO"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "HOLDER NAME"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Frame Frame2 
         Caption         =   "Brand"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   5775
         Left            =   -68160
         TabIndex        =   53
         Top             =   720
         Width           =   5655
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
            ItemData        =   "supplier_entry_form.frx":0054
            Left            =   2160
            List            =   "supplier_entry_form.frx":0056
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   600
            Width           =   2055
         End
         Begin VB.ListBox List3 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000009&
            Height          =   2370
            ItemData        =   "supplier_entry_form.frx":0058
            Left            =   3120
            List            =   "supplier_entry_form.frx":005A
            TabIndex        =   56
            Top             =   3000
            Width           =   1095
         End
         Begin VB.CommandButton Command2 
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
            Left            =   4440
            Style           =   1  'Graphical
            TabIndex        =   55
            Top             =   3360
            Width           =   975
         End
         Begin VB.ListBox List5 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   2370
            ItemData        =   "supplier_entry_form.frx":005C
            Left            =   1920
            List            =   "supplier_entry_form.frx":005E
            TabIndex        =   54
            Top             =   3000
            Width           =   1215
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
            ItemData        =   "supplier_entry_form.frx":0060
            Left            =   2160
            List            =   "supplier_entry_form.frx":0062
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   1320
            Width           =   2055
         End
         Begin VB.ListBox List6 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   2370
            ItemData        =   "supplier_entry_form.frx":0064
            Left            =   840
            List            =   "supplier_entry_form.frx":0066
            TabIndex        =   64
            Top             =   3000
            Width           =   1095
         End
         Begin VB.ListBox List14 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   1785
            Left            =   2040
            TabIndex        =   63
            Top             =   3360
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.ListBox List15 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   2370
            ItemData        =   "supplier_entry_form.frx":0068
            Left            =   840
            List            =   "supplier_entry_form.frx":006A
            TabIndex        =   62
            Top             =   3000
            Width           =   1095
         End
         Begin VB.TextBox Text4 
            Height          =   375
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   69
            Text            =   " "
            Top             =   600
            Width           =   2055
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   68
            Text            =   " "
            Top             =   1320
            Width           =   2055
         End
         Begin VB.Label Label48 
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
            TabIndex        =   93
            Top             =   1320
            Width           =   120
         End
         Begin VB.Label Label47 
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
            TabIndex        =   92
            Top             =   600
            Width           =   120
         End
         Begin VB.Line Line10 
            X1              =   3120
            X2              =   3120
            Y1              =   2760
            Y2              =   3000
         End
         Begin VB.Line Line8 
            X1              =   1920
            X2              =   1920
            Y1              =   2760
            Y2              =   3000
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
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
            Height          =   240
            Left            =   960
            TabIndex        =   61
            Top             =   600
            Width           =   630
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Unit"
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
            Left            =   3000
            TabIndex        =   60
            Top             =   2760
            Width           =   1095
         End
         Begin VB.Label Label28 
            Alignment       =   2  'Center
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
            Height          =   255
            Left            =   1920
            TabIndex        =   59
            Top             =   2760
            Width           =   1095
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " Ref"
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
            Left            =   1185
            TabIndex        =   58
            Top             =   2760
            Width           =   405
         End
         Begin VB.Shape Shape4 
            Height          =   255
            Left            =   840
            Top             =   2760
            Width           =   3375
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
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
            Left            =   1080
            TabIndex        =   57
            Top             =   1320
            Width           =   420
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H80000004&
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   5775
         Left            =   -74880
         TabIndex        =   39
         Top             =   720
         Width           =   6735
         Begin VB.ListBox List4 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000009&
            Height          =   2370
            ItemData        =   "supplier_entry_form.frx":006C
            Left            =   3840
            List            =   "supplier_entry_form.frx":006E
            TabIndex        =   52
            Top             =   3240
            Width           =   1095
         End
         Begin VB.ListBox List13 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000009&
            Height          =   2370
            ItemData        =   "supplier_entry_form.frx":0070
            Left            =   1560
            List            =   "supplier_entry_form.frx":0072
            TabIndex        =   51
            Top             =   3240
            Width           =   1095
         End
         Begin VB.ListBox List2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000009&
            Height          =   2370
            ItemData        =   "supplier_entry_form.frx":0074
            Left            =   480
            List            =   "supplier_entry_form.frx":0076
            TabIndex        =   50
            Top             =   3240
            Width           =   1095
         End
         Begin VB.ListBox List1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H8000000E&
            Height          =   2370
            ItemData        =   "supplier_entry_form.frx":0078
            Left            =   2640
            List            =   "supplier_entry_form.frx":007A
            TabIndex        =   42
            Top             =   3240
            Width           =   1215
         End
         Begin VB.TextBox gst 
            Height          =   375
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   13
            Text            =   " "
            Top             =   1800
            Width           =   2055
         End
         Begin VB.TextBox product_name 
            Height          =   375
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   12
            Text            =   " "
            Top             =   1200
            Width           =   2055
         End
         Begin VB.CommandButton delete_supplier_product 
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
            Left            =   5400
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   4440
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
            Left            =   5400
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   3480
            Width           =   1100
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
            ForeColor       =   &H00000000&
            Height          =   360
            ItemData        =   "supplier_entry_form.frx":007C
            Left            =   2880
            List            =   "supplier_entry_form.frx":007E
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   600
            Width           =   2055
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   67
            Text            =   " "
            Top             =   600
            Width           =   2055
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
            Left            =   2160
            TabIndex        =   91
            Top             =   1800
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
            Left            =   2400
            TabIndex        =   90
            Top             =   1200
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
            Left            =   2280
            TabIndex        =   89
            Top             =   600
            Width           =   120
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "GST"
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
            Left            =   3840
            TabIndex        =   49
            Top             =   3015
            Width           =   1095
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
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
            Left            =   2760
            TabIndex        =   48
            Top             =   3015
            Width           =   1095
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "ID"
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
            Left            =   1560
            TabIndex        =   47
            Top             =   3015
            Width           =   1095
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
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
            Height          =   255
            Left            =   480
            TabIndex        =   46
            Top             =   3015
            Width           =   1095
         End
         Begin VB.Line Line7 
            X1              =   3840
            X2              =   3840
            Y1              =   3000
            Y2              =   3240
         End
         Begin VB.Line Line6 
            X1              =   2640
            X2              =   2640
            Y1              =   3000
            Y2              =   3240
         End
         Begin VB.Line Line1 
            X1              =   1560
            X2              =   1560
            Y1              =   3000
            Y2              =   3240
         End
         Begin VB.Shape Shape3 
            Height          =   255
            Left            =   480
            Top             =   3000
            Width           =   4455
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   "GST %"
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
            Left            =   1320
            TabIndex        =   45
            Top             =   1800
            Width           =   735
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   " Product ID"
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
            Left            =   960
            TabIndex        =   44
            Top             =   600
            Width           =   1155
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
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
            Height          =   240
            Left            =   720
            TabIndex        =   43
            Top             =   1200
            Width           =   1545
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H80000004&
         Caption         =   "Account Info"
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
         Height          =   1815
         Left            =   -73560
         TabIndex        =   23
         Top             =   4320
         Width           =   9615
         Begin VB.ComboBox bank_name 
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
            ItemData        =   "supplier_entry_form.frx":0080
            Left            =   6720
            List            =   "supplier_entry_form.frx":0099
            TabIndex        =   9
            Top             =   480
            Width           =   2295
         End
         Begin VB.TextBox account_no 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   2400
            MaxLength       =   16
            TabIndex        =   7
            Top             =   480
            Width           =   2295
         End
         Begin VB.TextBox account_holder 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   2400
            MaxLength       =   30
            TabIndex        =   8
            Top             =   1080
            Width           =   2295
         End
         Begin VB.TextBox ifsc_code 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   6720
            MaxLength       =   16
            TabIndex        =   10
            Top             =   1080
            Width           =   2295
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
            Left            =   6360
            TabIndex        =   88
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
            Left            =   6480
            TabIndex        =   87
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
            Left            =   2160
            TabIndex        =   86
            Top             =   1200
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
            Left            =   1920
            TabIndex        =   85
            Top             =   600
            Width           =   120
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   " Account No"
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
            TabIndex        =   27
            Top             =   600
            Width           =   1260
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   " Account Holder"
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
            Left            =   360
            TabIndex        =   26
            Top             =   1200
            Width           =   1665
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   " Ifsc Code"
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
            Left            =   5160
            TabIndex        =   25
            Top             =   1080
            Width           =   1050
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   " Bank Name"
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
            Left            =   5040
            TabIndex        =   24
            Top             =   480
            Width           =   1275
         End
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2655
         Left            =   1800
         TabIndex        =   66
         Top             =   4320
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   4683
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
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
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000004&
         Caption         =   "Supplier Info"
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
         Height          =   2895
         Left            =   -73560
         TabIndex        =   28
         Top             =   960
         Width           =   9495
         Begin VB.TextBox supplier_name 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   2400
            MaxLength       =   30
            TabIndex        =   0
            Top             =   960
            Width           =   2415
         End
         Begin VB.TextBox supplier_idfe 
            Enabled         =   0   'False
            Height          =   375
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   360
            Width           =   2415
         End
         Begin VB.TextBox phone_no 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   6480
            MaxLength       =   10
            TabIndex        =   4
            Top             =   840
            Width           =   2535
         End
         Begin VB.TextBox email 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   6480
            MaxLength       =   30
            TabIndex        =   5
            Top             =   1440
            Width           =   2535
         End
         Begin VB.TextBox fax_no 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   2400
            MaxLength       =   16
            TabIndex        =   2
            Top             =   2160
            Width           =   2415
         End
         Begin VB.TextBox gstin_no 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   6480
            MaxLength       =   14
            TabIndex        =   3
            Top             =   240
            Width           =   2535
         End
         Begin VB.TextBox company_name 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   2400
            MaxLength       =   30
            TabIndex        =   1
            Top             =   1560
            Width           =   2415
         End
         Begin RichTextLib.RichTextBox address 
            Height          =   615
            Left            =   6480
            TabIndex        =   6
            Top             =   1920
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   1085
            _Version        =   393217
            BackColor       =   16777215
            MaxLength       =   100
            TextRTF         =   $"supplier_entry_form.frx":0104
         End
         Begin VB.ComboBox Combo2 
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
            TabIndex        =   30
            Text            =   "Select Supplier ID"
            Top             =   360
            Visible         =   0   'False
            Width           =   2415
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
            Left            =   6000
            TabIndex        =   84
            Top             =   2040
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
            Left            =   5760
            TabIndex        =   83
            Top             =   1560
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
            Left            =   6120
            TabIndex        =   82
            Top             =   960
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
            Left            =   6120
            TabIndex        =   81
            Top             =   360
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
            Left            =   1560
            TabIndex        =   80
            Top             =   2160
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
            Left            =   2160
            TabIndex        =   79
            Top             =   1680
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
            Left            =   2040
            TabIndex        =   78
            Top             =   1080
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
            TabIndex        =   77
            Top             =   480
            Width           =   120
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   " Supplier ID"
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
            TabIndex        =   38
            Top             =   480
            Width           =   1230
         End
         Begin VB.Label Label7 
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   " Phone No"
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
            Left            =   4920
            TabIndex        =   37
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label8 
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   " Email"
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
            Left            =   5040
            TabIndex        =   36
            Top             =   1560
            Width           =   1575
         End
         Begin VB.Label Label5 
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   " Address"
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
            Left            =   4920
            TabIndex        =   35
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   " Fax No"
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
            TabIndex        =   34
            Top             =   2160
            Width           =   810
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
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
            Height          =   240
            Left            =   5040
            TabIndex        =   33
            Top             =   360
            Width           =   960
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
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
            Height          =   240
            Left            =   360
            TabIndex        =   32
            Top             =   1680
            Width           =   1725
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackColor       =   &H80000002&
            BackStyle       =   0  'Transparent
            Caption         =   " Supplier Name"
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
            Left            =   360
            TabIndex        =   31
            Top             =   1080
            Width           =   1620
         End
      End
      Begin VB.Line Line2 
         BorderStyle     =   3  'Dot
         X1              =   240
         X2              =   12360
         Y1              =   3840
         Y2              =   3840
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
         Left            =   8520
         TabIndex        =   76
         Top             =   3960
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
         Left            =   4920
         TabIndex        =   75
         Top             =   3960
         Width           =   165
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   " Supplier Product Information"
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
         Left            =   5160
         TabIndex        =   74
         Top             =   3960
         Width           =   3285
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
         Left            =   8040
         TabIndex        =   73
         Top             =   600
         Width           =   165
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
         Left            =   5280
         TabIndex        =   72
         Top             =   600
         Width           =   165
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   " Supplier Information"
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
         TabIndex        =   71
         Top             =   600
         Width           =   2370
      End
      Begin VB.Image Image2 
         Height          =   420
         Left            =   -74760
         Picture         =   "supplier_entry_form.frx":0186
         Top             =   6840
         Width           =   1335
      End
      Begin VB.Image Image1 
         Height          =   420
         Left            =   -64080
         Picture         =   "supplier_entry_form.frx":0B92
         Top             =   6720
         Width           =   1335
      End
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
      Left            =   12600
      MaskColor       =   &H00000080&
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   8760
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
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   8760
      Width           =   1695
   End
   Begin VB.CommandButton new 
      Appearance      =   0  'Flat
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
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8760
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
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   8760
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
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   8760
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   14640
      TabIndex        =   21
      Top             =   2160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "SUPPLIER  DETAIL"
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
      Left            =   4560
      TabIndex        =   70
      Top             =   120
      Width           =   9045
   End
End
Attribute VB_Name = "supplier_entry_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim s As String
Dim i, j As Integer
Public m As String
Dim X As Integer
Dim item As ListItem
Public ind As Integer


Private Sub account_holder_GotFocus()
account_holder.Text = supplier_name.Text
End Sub



Private Sub account_holder_KeyPress(KeyAscii As Integer)
account_holder.BackColor = vbWhite
Select Case KeyAscii
 Case 33 To 64, 91 To 96, 123 To 126
  MsgBox "Must be a latter", vbCritical, "Warning"
  KeyAscii = 0
End Select
End Sub

Private Sub account_no_Change()
If account_no.Text = "" Then
 account_holder.Enabled = False
 ifsc_code.Enabled = False
 bank_name.Enabled = False
Else
  account_holder.Enabled = True
 ifsc_code.Enabled = True
 bank_name.Enabled = True
End If
End Sub

Private Sub account_no_KeyPress(KeyAscii As Integer)
account_no.BackColor = vbWhite
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn Then
  If KeyAscii = 13 Then
  End If
Else
KeyAscii = 0
a = MsgBox("Enter only number", vbCritical, "Warning")
End If
End Sub

Private Sub add_Click()
If Combo1.Text = "" Or Combo1.Text = "Select product ID" Or product_name.Text = "" Or gst.Text = "" Then
MsgBox "Some fields are Blank", vbCritical
Else
Combo5.clear

List1.AddItem product_name.Text
List4.AddItem gst.Text
List13.AddItem Combo1.Text
Dim a As String
Dim j, i As Integer
j = List2.ListCount
Set r = New ADODB.Recordset
sql = "select max(to_number(substr(s_no,5,length(s_no)))) from suppliersno"
Set r = c.Execute(sql)
If IsNull(r.Fields(0)) Then
List2.AddItem "sn" & "00" & 1
Else
List2.AddItem "sn" & "00" & r.Fields(0) + 1
End If
a = List2.List(j)
If (a = "SU" & "001" & "0") Then
sql = "select max(to_number(substr(s_no,4,length(s_no)))) from suppliersno"
Set r = c.Execute(sql)
List2.List(j) = "sn" & "0" & r.Fields(0) + 1
End If
Text2.Text = List2.List(j)
Set r = New ADODB.Recordset
sql = "insert into suppliersno values('" + List2.List(j) + "')"
Set r = c.Execute(sql)


sql = "select DISTINCT(BRAND) from product_brand where product_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)
Combo4.clear
Do While Not r.EOF
Combo4.AddItem r!brand
r.MoveNext
Loop
List14.AddItem Combo1.Text
i = Combo1.ListIndex
Combo1.RemoveItem i
End If
product_name.Text = ""
gst.Text = ""
End Sub



Private Sub addacc1_Click()
List7.AddItem account_no.Text
List8.AddItem account_holder.Text
List9.AddItem ifsc_code.Text
List10.AddItem bank_name.Text
List11.AddItem branch.Text
Dim a As String
Dim j As Integer
j = List12.ListCount
Set r = New ADODB.Recordset
sql = "select max(to_number(substr(s_no,5,length(s_no)))) from supplieraccountsno"
Set r = c.Execute(sql)
If IsNull(r.Fields(0)) Then
List12.AddItem "sn" & "00" & 1
Else
List12.AddItem "sn" & "00" & r.Fields(0) + 1
End If
a = List12.List(j)
If (a = "SU" & "001" & "0") Then
sql = "select max(to_number(substr(s_no,4,length(s_no)))) from supplieraccountsno"
Set r = c.Execute(sql)
List12.List(j) = "sn" & "0" & r.Fields(0) + 1
End If
Set r = New ADODB.Recordset
sql = "insert into supplieraccountsno values('" & List12.List(j) & "')"
Set r = c.Execute(sql)
End Sub





Private Sub address_KeyPress(KeyAscii As Integer)
address.BackColor = vbWhite
End Sub

Private Sub bank_name_KeyPress(KeyAscii As Integer)
bank_name.BackColor = vbWhite
Select Case KeyAscii
 Case 33 To 64, 91 To 96, 123 To 126
  MsgBox "Must be a latter", vbCritical, "Warning"
  KeyAscii = 0
End Select
End Sub

Private Sub bank_name_LostFocus()
bank_name.Text = UCase(bank_name.Text)
End Sub

Private Sub Combo1_Click()
additemcombo
End Sub

Private Sub Combo2_Change()
If Combo1.Text <> blank Then
Combo4.clear
sql = "select DISTINCT(BRAND) from product_brand where product_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)
Do While Not r.EOF
Combo4.AddItem r!brand
r.MoveNext
Loop
End If
End Sub

Private Sub Combo2_Click()
Set r = New ADODB.Recordset
sql = "select *from supplier_detail where supplier_id='" + Combo2.Text + "'"
Set r = c.Execute(sql)

supplier_name.Text = r.Fields(7)
company_name.Text = r.Fields(3)
fax_no.Text = r.Fields(5)
gstin_no.Text = r.Fields(4)
phone_no.Text = r.Fields(1)
email.Text = r.Fields(2)
address.Text = r.Fields(6)


Set r = c.Execute("select * from supplier_account where supplier_id='" + Combo2.Text + "'")
account_no.Text = r.Fields(2)
account_holder.Text = r.Fields(3)
ifsc_code.Text = r.Fields(4)
bank_name.Text = r.Fields(5)

Set r = c.Execute("select * from supplier_product where s_id='" + Combo2.Text + "'")
Do While Not r.EOF
List2.AddItem r.Fields(0)
List13.AddItem r.Fields(2)
List1.AddItem r.Fields(3)

List4.AddItem r.Fields(4)
r.MoveNext
Loop



Set r = c.Execute("select * from supplier_product_brand where sup_id='" + Combo2.Text + "'")
Do While Not r.EOF

List6.AddItem r.Fields(3)
List5.AddItem r.Fields(0)
List3.AddItem r.Fields(4)
r.MoveNext
Loop
List15.Visible = False
List6.Visible = True
End Sub

Private Sub Combo4_click()
Combo5.clear
sql = "select (unit) from product_brand where brand='" + Combo4.Text + "'"
Set r = c.Execute(sql)
While r.EOF = False
Combo5.AddItem r.Fields(0)
r.MoveNext
Wend
End Sub

Private Sub Combo5_Click()
Dim X As Integer

X = List13.ListCount - 1
List3.AddItem Combo5.Text
List6.AddItem Text2.Text
sql = "select brand from product_brand where unit='" + Combo5.Text + "' and product_id='" + List13.List(X) + "'"
Set r = c.Execute(sql)
List5.AddItem r.Fields(0)
i = Combo5.ListIndex
Combo5.RemoveItem i
List15.AddItem j + 1
j = j + 1

End Sub

Private Sub account_holder_LostFocus()
If account_holder.Text = "" Then
Else
account_holder.Text = UCase(account_holder.Text)
End If

End Sub




Private Sub Command2_Click()
Dim i As Integer
Dim item As String
If List3.ListIndex = -1 Then
Else
item = List3.List(i)
Combo5.AddItem item
i = List3.ListIndex
List5.RemoveItem ind
List6.RemoveItem ind
List3.RemoveItem ind
i = i + 1
End If
End Sub

Private Sub company_name_KeyPress(KeyAscii As Integer)
company_name.BackColor = vbWhite
Select Case KeyAscii
 Case 33 To 64, 91 To 96, 123 To 126
  MsgBox "Must be a latter", vbCritical, "Warning"
  KeyAscii = 0
End Select
End Sub

Private Sub delete_Click()
ans = MsgBox("Do you Want to Delete All Supplier Detail ?", vbOKCancel + vbInformation)
If ans = 1 Then

Set r = c.Execute("delete supplier_product_brand where sup_id='" + Combo2.Text + "'")
Set r = c.Execute("delete supplier_product where s_id='" + Combo2.Text + "'")
Set r = c.Execute("delete supplier_account where supplier_id='" + Combo2.Text + "'")
Set r = c.Execute("delete supplier_detail where supplier_id='" + Combo2.Text + "'")
MsgBox "Deleted"
clear
End If
End Sub

Private Sub delete_supplier_product_Click()
Dim i As Integer
Dim item As String
i = 0
item = List13.List(i)
Combo1.AddItem item
i = i + 1

List1.RemoveItem ind
List2.RemoveItem ind
List13.RemoveItem ind
List4.RemoveItem ind
i = 0


End Sub

Private Sub deleteacc_Click()
Dim a, b, c, d, e, f As Integer
a = List7.ListIndex
b = List8.ListIndex
c = List9.ListIndex
d = List10.ListIndex
e = List11.ListIndex
If (a >= 0) And (b >= 0) And (c >= 0) And (d >= 0) And (d >= 0) And (e >= 0) Then
List7.RemoveItem a
List8.RemoveItem b
List9.RemoveItem c
List10.RemoveItem d
List11.RemoveItem e
End If

End Sub

Private Sub company_name_LostFocus()
If company_name.Text = "" Then
Else
 company_name.Text = UCase(company_name.Text)
End If

End Sub
Private Sub dist_LostFocus()
If dist.Text = "" Then
MsgBox "some foelds are empty"
dist.SetFocus
End If

End Sub


Private Sub first_name_LostFocus()

If first_name = "" Then
MsgBox "field is blank"
first_name.SetFocus
End If
End Sub


Private Sub email_KeyPress(KeyAscii As Integer)
email.BackColor = vbWhite
End Sub

Private Sub fax_no_KeyPress(KeyAscii As Integer)
fax_no.BackColor = vbWhite
End Sub


Private Sub Form_Load()
Connection
auto
supplier_idfe.Enabled = True

supplier_entry_form.Caption = "Supplier Entry"

sql = "select (product_id) from product_detail"
Set r = c.Execute(sql)
Combo1.clear
Do While Not r.EOF
Combo1.AddItem r!product_id
r.MoveNext
Loop

MDIForm1.Picture2.Visible = True

'for add all product in view tab
  view_sup
  
  SSTab1.Tab = 0
End Sub






Private Sub gstin_no_KeyPress(KeyAscii As Integer)
gstin_no.BackColor = vbWhite
End Sub

Private Sub ifsc_code_KeyPress(KeyAscii As Integer)
ifsc_code.BackColor = vbWhite
End Sub

Private Sub ifsc_code_LostFocus()

If ifsc_code.Text = "" Then
End If
ifsc_code.Text = UCase(ifsc_code.Text)
End Sub

Private Sub Image1_Click()
SSTab1.Tab = 1

End Sub

Private Sub Image2_Click()
SSTab1.Tab = 0
End Sub

Private Sub insert_Click()
ans = MsgBox("Do you Want to Save", vbOKCancel + vbInformation, "Save")
If ans = 1 Then

If supplier_idfe.Text = "" Then
ElseIf supplier_name.Text = "" Then
 supplier_name.BackColor = &HC0C0FF
 MsgBox "Supplier name Empty", vbCritical
ElseIf company_name.Text = "" Then
 company_name.BackColor = &HC0C0FF
 MsgBox "Company name Empty", vbCritical
ElseIf fax_no.Text = "" Then
 fax_no.BackColor = &HC0C0FF
 MsgBox "Fax no Empty", vbCritical
ElseIf gstin_no.Text = "" Then
 gstin_no.BackColor = &HC0C0FF
 MsgBox "Gst no Empty", vbCritical
ElseIf phone_no.Text = "" Then
 phone_no.BackColor = &HC0C0FF
 MsgBox "Phone no Empty", vbCritical
ElseIf email.Text = "" Then
 email.BackColor = &HC0C0FF
 MsgBox "Email fields is Empty", vbCritical
ElseIf address.Text = "" Then
 address.BackColor = &HC0C0FF
 MsgBox "Address fields is Empty", vbCritical
ElseIf account_no.Text = "" Then
 account_no.BackColor = &HC0C0FF
 MsgBox "Account no fields is Empty", vbCritical
ElseIf account_holder.Text = "" Then
 account_holder.BackColor = &HC0C0FF
 MsgBox "Account holder Name fields is Empty", vbCritical
ElseIf bank_name.Text = "" Then
 bank_name.BackColor = &HC0C0FF
 MsgBox "Bank name fields is Empty", vbCritical
ElseIf ifsc_code.Text = "" Then
 ifsc_code.BackColor = &HC0C0FF
 MsgBox "ifsc code fields is Empty", vbCritical
ElseIf List1.List(0) = "" Or List2.List(0) = "" Or List3.List(0) = "" Or List4.List(0) = "" Or List5.List(0) = "" Or List6.List(0) = "" Or List13.List(0) = "" Then
a = MsgBox("Some Fields are required", vbOKOnly + vbCritical, "warnnig")
Else
sql = "insert into supplier_detail values('" + supplier_idfe.Text + "','" + phone_no.Text + "','" + email.Text + "','" + company_name.Text + "','" + gstin_no.Text + "','" + fax_no.Text + "','" + address.Text + "','" + supplier_name.Text + "')"
Set r = c.Execute(sql)

For i = 0 To List1.ListCount - 1 Step 1
sql = "insert into supplier_product values ('" & List2.List(i) & "','" & supplier_idfe.Text & "','" & List13.List(i) & "','" & List1.List(i) & "','" & List4.List(i) & "')"

Set r = c.Execute(sql)
Next i

Set r = c.Execute("select max(s_no) from supplier_account")
If IsNull(r.Fields(0)) Then
gh = 0
Else
gh = r.Fields(0)
End If
sql = "insert into supplier_account values ('" & gh + 1 & "','" & supplier_idfe.Text & "','" & account_no.Text & "','" & account_holder.Text & "','" & ifsc_code.Text & "','" & bank_name.Text & "')"
Set r = c.Execute(sql)


For i = 0 To List5.ListCount - 1
sql = "insert into supplier_product_brand values('" & List5.List(i) & "','" & List14.List(i) & "','" & supplier_idfe.Text & "','" & List6.List(i) & "','" + List3.List(i) + "')"
Set r = c.Execute(sql)
Next i

a = MsgBox("Supplier  Saved", , "Save")
insert.Enabled = False
clear
auto
End If
End If
End Sub





Private Sub landmark_LostFocus()
If landmark.Text = "" Then
a = MsgBox("Email field is empty", vbOKOnly + vbCritical, "warnning")
landmark.SetFocus
End If

End Sub

Private Sub List1_Click()
ind = List1.ListIndex
i = List1.ListIndex
Text1.Text = List13.List(i)
product_name.Text = List1.List(i)
gst.Text = List4.List(i)
End Sub

Private Sub List13_Click()
ind = List13.ListIndex
i = List13.ListIndex
Text1.Text = List13.List(i)
product_name.Text = List1.List(i)
gst.Text = List4.List(i)
End Sub

Private Sub List2_Click()
ind = List2.ListIndex
i = List2.ListIndex
Text1.Text = List13.List(i)
product_name.Text = List1.List(i)
gst.Text = List4.List(i)
End Sub

Private Sub List3_Click()
i = List3.ListIndex
ind = List3.ListIndex
Text4.Text = List5.List(i)
Text3.Text = List3.List(i)
End Sub

Private Sub List4_Click()
ind = List4.ListIndex
i = List4.ListIndex
Text1.Text = List13.List(i)
product_name.Text = List1.List(i)
gst.Text = List4.List(i)
End Sub

Private Sub List5_Click()
i = List5.ListIndex
ind = List5.ListIndex
Text4.Text = List5.List(i)
Text3.Text = List3.List(i)
End Sub

Private Sub List6_Click()
i = List6.ListIndex
ind = List6.ListIndex
Text4.Text = List5.List(i)
Text3.Text = List3.List(i)
End Sub

Private Sub new_Click()
insert.Enabled = True
clear
auto
sql = "select product_id from product_detail"
Set r = c.Execute(sql)
Do While Not r.EOF
Combo1.AddItem r!product_id
r.MoveNext
Loop
Combo2.Visible = False
supplier_idfe.Visible = True
j = 0

insert.Enabled = True
update.Enabled = False
delete.Enabled = False

List15.Visible = True
List6.Visible = False

Combo1.Visible = True
Combo4.Visible = True
Combo5.Visible = True
Text3.Visible = False
Text4.Visible = False
Text1.Visible = False

sql = "select (product_id) from product_detail"
Set r = c.Execute(sql)
Combo1.clear
Do While Not r.EOF
Combo1.AddItem r!product_id
r.MoveNext
Loop
End Sub




Private Sub product_name_LostFocus()
product_name.Text = UCase(product_name.Text)
End Sub

Private Sub supplier_name_KeyPress(KeyAscii As Integer)
supplier_name.BackColor = vbWhite
Select Case KeyAscii
 Case 33 To 64, 91 To 96, 123 To 126
  MsgBox "Must be a latter", vbCritical, "Warning"
  KeyAscii = 0
End Select
End Sub

Private Sub supplier_name_LostFocus()
 supplier_name.Text = UCase(supplier_name.Text)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    sur_name.SetFocus
    End If
End Sub
Private Sub Label26_Click()
Unload customer_entry_form
Unload supplier_entry_form

Load product_detail
product_detail.Show
End Sub

Private Sub Label27_Click()
Unload customer_entry_form
Unload product_detail

Load supplier_entry_form
supplier_entry_form.Show
End Sub

Private Sub Label9_Click()
Load customer_entry_form
customer_entry_form.Show
End Sub


Public Function check()
If Option1.Value = True Then
s = "male"
Else
s = "female"
End If

End Function

Private Sub phone_no_KeyPress(KeyAscii As Integer)
phone_no.BackColor = vbWhite
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn Then
  If KeyAscii = 13 Then
  End If
Else
KeyAscii = 0
a = MsgBox("Enter only number", vbCritical, "Warning")
End If
End Sub

Private Sub phone_no_LostFocus()
If phone_no.Text = "" Then
ElseIf Len(phone_no.Text) <> 10 Then
 a = MsgBox("Phone no length not = 10", vbCritical, "Warning")
End If

End Sub



Private Sub state_lostfocus()
If State.Text = "" Then
MsgBox "some fields are empty"
State.SetFocus
End If
End Sub



Private Sub report_Click()
update.Enabled = True
delete.Enabled = True
insert.Enabled = False

Combo1.Visible = False
Combo4.Visible = False
Combo5.Visible = False
Text3.Visible = True
Text4.Visible = True
Text1.Visible = True

Combo2.Visible = True
supplier_idfe.Visible = False
SSTab1.Tab = 0

Combo2.clear
Set r = c.Execute("select supplier_id from supplier_detail")
Do While Not r.EOF
Combo2.AddItem r!supplier_id
r.MoveNext
Loop
insert.Enabled = False
End Sub

Private Sub sur_name_LostFocus()
If sur_name.Text = "" Then
MsgBox "surname field is empty"
sur_name.SetFocus
End If

End Sub

Public Function additemcombo()
Set r = New ADODB.Recordset
sql = "select * from product_detail where product_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)

product_name.Text = r.Fields(1)
gst.Text = r.Fields(2)

End Function




Private Sub SSTab1_GotFocus()
view_sup
End Sub

Private Sub update_Click()
Dim i As Integer
ans = MsgBox("Do you Want to Update All detail Of Supplier", vbOKCancel + vbInformation)
If ans = 1 Then
sql = "update supplier_detail set phone_no='" + phone_no.Text + "',email='" + email.Text + "',company_name='" + company_name.Text + "',gstin_no='" + gstin_no.Text + "',fax_no='" + fax_no.Text + "',address='" + address.Text + "',supplier_name='" + supplier_name.Text + "' where supplier_id='" + Combo2.Text + "' "

Set r = c.Execute(sql)


sql = "update supplier_account set account_no='" + account_no.Text + "',account_holder_name='" + account_holder.Text + "',ifsc_code='" + ifsc_code.Text + "',bank_name='" + bank_name + "' where supplier_id='" + Combo2.Text + "'"
Set r = c.Execute(sql)


For i = 0 To List1.ListCount - 1
sql = "update supplier_product set product_name='" + List1.List(i) + "',gst='" + List4.List(i) + "'where s_id='" + Combo2.Text + "'"
Set r = c.Execute(sql)
Next i

For i = 0 To List1.ListCount - 1
sql = "update supplier_product_brand set brand='" + List5.List(i) + "',unit='" + List3.List(i) + "'where sup_id='" + Combo2.Text + "'"
Set r = c.Execute(sql)
Next i
MsgBox "Update Completed"
clear
End If
End Sub

Private Sub updateacc_Click()
List7.clear
List8.clear
List9.clear
List10.clear
List11.clear
List12.clear

Set r = New ADODB.Recordset
sql = "select (account_no) from supplier_account "
Set r = c.Execute(sql)
Do While Not r.EOF
Combo4.AddItem r!account_no
r.MoveNext
Loop

End Sub



Private Sub view_Click()
Load supplier_form_view
supplier_form_view.Show

End Sub

Public Function clear()
supplier_idfe.Text = ""
supplier_name.Text = ""
company_name.Text = ""
gstin_no.Text = ""
fax_no.Text = ""
phone_no.Text = ""
email.Text = ""
address.Text = ""
account_no.Text = ""
account_holder.Text = ""
ifsc_code.Text = ""
bank_name.Text = ""

product_name.Text = ""
'unit.Text = ""
gst.Text = ""
List2.clear
List13.clear
List1.clear
List3.clear
List4.clear
'List14.clear
List6.clear
List5.clear
List15.clear
Combo1.clear
Combo4.clear
Combo5.clear

End Function



Public Function auto()

Dim a As String

sql = "select max(to_number(substr(supplier_id,5,length(supplier_id)))) from supplier_detail"
Set r = c.Execute(sql)
If IsNull(r.Fields(0)) Then
supplier_idfe.Text = "SU" & "00" & 1
Else
supplier_idfe.Text = "SU" & "00" & r.Fields(0) + 1
End If
a = supplier_idfe.Text
If (a = "SU" & "001" & "0") Then
sql = "select max(to_number(substr(supplier_id,4,length(supplier_id)))) from supplier_detail"
Set r = c.Execute(sql)
supplier_idfe.Text = "SU" & "0" & r.Fields(0) + 1
End If
End Function

Public Function view_sup()
Set r = c.Execute("select * from supplier_detail,supplier_account")
ListView1.ListItems.clear
ListView2.ListItems.clear
While Not r.EOF

Set item = ListView1.ListItems.add(, , r.Fields(0))
item.SubItems(1) = r.Fields(7)
item.SubItems(2) = r.Fields(3)
item.SubItems(3) = r.Fields(1)
item.SubItems(4) = r.Fields(6)
item.SubItems(5) = r.Fields(4)
item.SubItems(6) = r.Fields(13)
item.SubItems(7) = r.Fields(12)
item.SubItems(8) = r.Fields(10)
item.SubItems(9) = r.Fields(11)

r.MoveNext
Wend


Set r = c.Execute("select * from supplier_product_brand,supplier_product")
While Not r.EOF

Set item = ListView2.ListItems.add(, , r.Fields(3))
item.SubItems(1) = r.Fields(7)
item.SubItems(2) = r.Fields(8)
item.SubItems(3) = r.Fields(0)
item.SubItems(4) = r.Fields(4)
item.SubItems(5) = r.Fields(9)
r.MoveNext
Wend
End Function


