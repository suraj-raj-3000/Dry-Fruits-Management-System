VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form customer_entry_form 
   BackColor       =   &H80000004&
   Caption         =   "Form1"
   ClientHeight    =   8355
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15840
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8355
   ScaleWidth      =   15840
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton savecustomer 
      BackColor       =   &H0080FF80&
      Caption         =   " SAVE"
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
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8640
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
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8640
      Width           =   1695
   End
   Begin VB.CommandButton newproduct 
      BackColor       =   &H00C0C0C0&
      Caption         =   "NEW"
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
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   7
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
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8640
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
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8640
      Width           =   1695
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7575
      Left            =   2280
      TabIndex        =   12
      Top             =   600
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   13361
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
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
      TabCaption(0)   =   "CUSTOMER"
      TabPicture(0)   =   "costumer_entry_form.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "VIEW"
      TabPicture(1)   =   "costumer_entry_form.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label13"
      Tab(1).Control(1)=   "Label25"
      Tab(1).Control(2)=   "Label26"
      Tab(1).Control(3)=   "ListView1"
      Tab(1).ControlCount=   4
      Begin VB.Frame Frame3 
         Caption         =   "-:  Address  :-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   2895
         Left            =   6840
         TabIndex        =   36
         Top             =   1080
         Width           =   4695
         Begin VB.TextBox address 
            Height          =   375
            Left            =   1680
            MaxLength       =   55
            MultiLine       =   -1  'True
            TabIndex        =   3
            Text            =   "costumer_entry_form.frx":0038
            Top             =   480
            Width           =   2775
         End
         Begin VB.ComboBox combo2 
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
            ItemData        =   "costumer_entry_form.frx":003A
            Left            =   1680
            List            =   "costumer_entry_form.frx":003C
            TabIndex        =   5
            Top             =   1680
            Width           =   2775
         End
         Begin VB.ComboBox combo1 
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
            ItemData        =   "costumer_entry_form.frx":003E
            Left            =   1680
            List            =   "costumer_entry_form.frx":00AE
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1080
            Width           =   2775
         End
         Begin VB.TextBox Text6 
            Height          =   375
            Left            =   1680
            MaxLength       =   30
            TabIndex        =   6
            Text            =   " "
            Top             =   2160
            Width           =   2775
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
            Left            =   1080
            TabIndex        =   50
            Top             =   2280
            Width           =   120
         End
         Begin VB.Label Label19 
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
            Left            =   1320
            TabIndex        =   49
            Top             =   1680
            Width           =   120
         End
         Begin VB.Label Label18 
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
            Left            =   1080
            TabIndex        =   48
            Top             =   1200
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
            Left            =   1440
            TabIndex        =   47
            Top             =   600
            Width           =   120
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   360
            TabIndex        =   40
            Top             =   600
            Width           =   945
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " District"
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
            TabIndex        =   39
            Top             =   1680
            Width           =   795
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " State"
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
            TabIndex        =   38
            Top             =   1200
            Width           =   615
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   360
            TabIndex        =   37
            Top             =   2280
            Width           =   660
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   6015
         Left            =   -74880
         TabIndex        =   34
         Top             =   960
         Width           =   13575
         _ExtentX        =   23945
         _ExtentY        =   10610
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
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "NAME"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "PHONE"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "ADDRESS"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "GSTIN"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "BANK"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "IFSC CODE"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "ACCOUNT NO"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Text            =   "HOLDER NAME"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000004&
         Caption         =   "-: Customer Info :-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   2895
         Left            =   2040
         TabIndex        =   13
         Top             =   1080
         Width           =   4815
         Begin VB.ComboBox Combo3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   480
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.TextBox Text4 
            Height          =   375
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   2
            Text            =   " "
            Top             =   2280
            Width           =   2415
         End
         Begin VB.CommandButton Command10 
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
            Height          =   675
            Left            =   6360
            TabIndex        =   18
            Top             =   6240
            Width           =   1455
         End
         Begin VB.CommandButton Command7 
            Caption         =   " UPDATE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Left            =   4560
            TabIndex        =   17
            Top             =   6240
            Width           =   1455
         End
         Begin VB.CommandButton Command4 
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
            Height          =   675
            Left            =   1440
            TabIndex        =   16
            Top             =   6240
            Width           =   1215
         End
         Begin VB.CommandButton Command3 
            Caption         =   " SAVE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Left            =   3000
            TabIndex        =   15
            Top             =   6240
            Width           =   1215
         End
         Begin VB.CommandButton Command2 
            Caption         =   " EXIT"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Left            =   8160
            TabIndex        =   14
            Top             =   6240
            Width           =   1455
         End
         Begin VB.TextBox gstin 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   2160
            MaxLength       =   16
            TabIndex        =   1
            Text            =   " "
            Top             =   1680
            Width           =   2415
         End
         Begin VB.TextBox customer_name 
            Height          =   375
            Left            =   2160
            MaxLength       =   30
            TabIndex        =   0
            Text            =   " "
            Top             =   1080
            Width           =   2415
         End
         Begin VB.TextBox customer_id 
            Enabled         =   0   'False
            Height          =   375
            Left            =   2160
            TabIndex        =   20
            Text            =   " "
            Top             =   480
            Width           =   2415
         End
         Begin VB.Label Label16 
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
            TabIndex        =   46
            Top             =   2400
            Width           =   120
         End
         Begin VB.Label Label14 
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
            TabIndex        =   45
            Top             =   1200
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
            Left            =   1680
            TabIndex        =   44
            Top             =   600
            Width           =   120
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   360
            TabIndex        =   24
            Top             =   2400
            Width           =   1095
         End
         Begin VB.Image Image4 
            Height          =   750
            Left            =   9840
            Picture         =   "costumer_entry_form.frx":0264
            Top             =   6240
            Width           =   750
         End
         Begin VB.Image Image3 
            Height          =   750
            Left            =   360
            Picture         =   "costumer_entry_form.frx":45D2
            Top             =   6240
            Width           =   750
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
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
            Left            =   360
            TabIndex        =   23
            Top             =   1800
            Width           =   960
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
            Left            =   240
            TabIndex        =   22
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " Customer Name"
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
            Left            =   120
            TabIndex        =   21
            Top             =   1200
            Width           =   1725
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000004&
         Caption         =   "-: Account Detail :-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   2895
         Left            =   4080
         TabIndex        =   25
         Top             =   4080
         Width           =   5295
         Begin VB.TextBox bank_name 
            Height          =   375
            Left            =   2280
            MaxLength       =   23
            TabIndex        =   29
            Text            =   " "
            Top             =   2160
            Width           =   2655
         End
         Begin VB.TextBox ifsc_code 
            Height          =   375
            Left            =   2280
            MaxLength       =   12
            TabIndex        =   28
            Text            =   " "
            Top             =   1560
            Width           =   2655
         End
         Begin VB.TextBox holder_name 
            Height          =   375
            Left            =   2280
            TabIndex        =   27
            Text            =   " "
            Top             =   960
            Width           =   2655
         End
         Begin VB.TextBox account_no 
            Height          =   375
            Left            =   2280
            MaxLength       =   12
            TabIndex        =   26
            Text            =   " "
            Top             =   360
            Width           =   2655
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
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
            Left            =   480
            TabIndex        =   33
            Top             =   2280
            Width           =   1275
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
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
            Left            =   480
            TabIndex        =   32
            Top             =   1680
            Width           =   1050
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " Holder Name"
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
            TabIndex        =   31
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            Left            =   480
            TabIndex        =   30
            Top             =   480
            Width           =   1260
         End
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
         Left            =   -66720
         TabIndex        =   43
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
         Left            =   -69480
         TabIndex        =   42
         Top             =   600
         Width           =   165
      End
      Begin VB.Label Label13 
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
         Left            =   -69240
         TabIndex        =   41
         Top             =   600
         Width           =   2370
      End
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "CUSTOMER DETAIL"
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
      Left            =   5040
      TabIndex        =   35
      Top             =   120
      Width           =   9045
   End
End
Attribute VB_Name = "customer_entry_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c1 As ADODB.Connection
Dim r1 As ADODB.Recordset
Dim sql As String
Dim item As ListItem





Private Sub account_no_Change()
If account_no.Text = "" Then
 holder_name.Enabled = False
 ifsc_code.Enabled = False
 bank_name.Enabled = False
Else
  holder_name.Enabled = True
 ifsc_code.Enabled = True
 bank_name.Enabled = True
End If
End Sub

Private Sub account_no_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn Then
  If KeyAscii = 13 Then
   account_no.SetFocus
  End If
Else
KeyAscii = 0
MsgBox "Enter only number"
End If
End Sub





Private Sub address_KeyPress(KeyAscii As Integer)
address.BackColor = vbWhite
End Sub

Private Sub bank_name_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
 Case 33 To 64, 91 To 96, 123 To 126
  MsgBox "Must be a latter", vbCritical, "Warning"
  KeyAscii = 0
End Select
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
Combo1.BackColor = vbWhite
End Sub



Private Sub Combo2_KeyPress(KeyAscii As Integer)
Combo2.BackColor = vbWhite
End Sub

Private Sub Combo3_Click()
additem_combo
End Sub



Private Sub customer_id_LostFocus()

If customer_id.Text = "" Then
MsgBox "customer id empty"
customer_id.SetFocus
End If
End Sub

Private Sub customer_name_KeyPress(KeyAscii As Integer)
customer_name.BackColor = vbWhite
Select Case KeyAscii
 Case 33 To 64, 91 To 96, 123 To 126
  MsgBox "Must be a latter", vbCritical, "Warning"
  KeyAscii = 0
End Select

End Sub

Private Sub customer_name_LostFocus()
If customer_name.Text = "" Then
Else
  customer_name.Text = UCase(customer_name.Text)
End If
End Sub

Private Sub delete_Click()
ans = MsgBox("Do you Want to Delete", vbOKCancel + vbInformation)
If ans = 1 Then
Set r1 = New ADODB.Recordset
sql = "delete from customer_detail where customer_id='" + Combo3.Text + "'"
Set r1 = c1.Execute(sql)
MsgBox "record deleted"
auto
Combo3.clear
combo_item
Adodc1.Refresh
End If
End Sub

Private Sub Form_Load()

Set c1 = New ADODB.Connection
c1.Open "Provider=MSDAORA.1;Password=lnt123;User ID=LNT;Persist Security Info=True"
auto
combo_item
MDIForm1.Picture2.Visible = True
customer_entry_form.Caption = "Customer Form"

ListView1.ListItems.clear
Set r1 = c1.Execute("select * from customer_detail")

If IsNull(r1.Fields(11)) Then
 rf11 = None
Else
 rf11 = r1.Fields(11)
End If

If IsNull(r1.Fields(9)) Then
 rf9 = None
Else
 rf9 = r1.Fields(9)
End If

If IsNull(r1.Fields(8)) Then
 rf8 = None
Else
 rf8 = r1.Fields(8)
End If

If IsNull(r1.Fields(10)) Then
 rf10 = None
Else
 rf10 = r1.Fields(10)
End If
While Not r1.EOF
Set item = ListView1.ListItems.add(, , r1.Fields(0))
item.SubItems(1) = r1.Fields(1)
item.SubItems(2) = r1.Fields(6)
item.SubItems(3) = r1.Fields(3)
item.SubItems(4) = r1.Fields(2)


item.SubItems(5) = rf11
item.SubItems(6) = rf10
item.SubItems(7) = rf8
item.SubItems(8) = rf9
r1.MoveNext
Wend
SSTab1.Tab = 0
End Sub

Private Sub gstin_KeyPress(KeyAscii As Integer)
gstin.BackColor = vbWhite
End Sub

Private Sub gstin_LostFocus()
If gstin.Text = "" Then
Else
gstin.Text = UCase(gstin.Text)
End If
End Sub




Private Sub holder_name_GotFocus()
holder_name.Text = customer_name.Text
End Sub

Private Sub holder_name_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
 Case 33 To 64, 91 To 96, 123 To 126
  MsgBox "Must be a latter", vbCritical, "Warning"
  KeyAscii = 0
End Select
End Sub

Private Sub ifsc_code_LostFocus()
ifsc_code.Text = UCase(ifsc_code.Text)
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
Unload product_detail
Unload supplier_entry_form

Load customer_entry_form
customer_entry_form.Show
End Sub

Private Sub newproduct_Click()
auto
Combo3.Visible = False
update.Enabled = False
delete.Enabled = False
savecustomer.Enabled = True

End Sub

Private Sub savecustomer_Click()
ans = MsgBox("Do you Want to Save", vbOKCancel + vbInformation)
If ans = 1 Then

If customer_name.Text = "" Then
 customer_name.BackColor = &HC0C0FF
 MsgBox "Customer name fields is Empty", vbCritical
ElseIf gstin.Text = "" Then
 gstin.BackColor = &HC0C0FF
  MsgBox "Gstin fields is Empty", vbCritical
ElseIf Text4.Text = "" Then
 Text4.BackColor = &HC0C0FF
 MsgBox "Phone no fields is Empty", vbCritical
ElseIf address.Text = "" Then
 address.BackColor = &HC0C0FF
 MsgBox "Address fields is Empty", vbCritical
ElseIf Combo1.Text = "" Then
 Combo1.BackColor = &HC0C0FF
 MsgBox "State fields is Empty", vbCritical
ElseIf Combo2.Text = "" Then
 Combo2.BackColor = &HC0C0FF
 MsgBox "District fields is Empty", vbCritical
ElseIf Text6.Text = "" Then
 Text6.BackColor = &HC0C0FF
 MsgBox "Email fields is Empty", vbCritical
MsgBox ""
Else
Set r1 = New ADODB.Recordset
sql = "insert into  customer_detail values('" + customer_id.Text + "','" + customer_name.Text + "'," + gstin.Text + ",'" + address.Text + "','" + Combo1.Text + "','" + Combo2.Text + "','" + Text4.Text + "','" + Text6.Text + "','" + account_no.Text + "','" + holder_name.Text + "','" + ifsc_code.Text + "','" + bank_name.Text + "')"
Set r1 = c1.Execute(sql)
MsgBox "data saved"
auto

Combo3.clear
combo_item
End If
End If
End Sub


Public Function auto()
customer_name.Text = ""
gstin.Text = ""
address.Text = ""

Combo2.Text = ""
Text4.Text = ""
Text6.Text = ""
account_no.Text = ""
holder_name.Text = ""
ifsc_code.Text = ""
bank_name.Text = ""
Set r1 = New ADODB.Recordset
sql = "select count(customer_id) from customer_detail"
Set r1 = c1.Execute(sql)
customer_id = r1.Fields(0) + 1
customer_id = "cust" + customer_id

End Function
Public Function combo_item()
Set r1 = New ADODB.Recordset
sql = "select customer_id from customer_detail "
Set r1 = c1.Execute(sql)
Do While Not r1.EOF
Combo3.AddItem r1!customer_id
r1.MoveNext
Loop
End Function




Private Sub search_Click()
Combo3.Visible = True
update.Enabled = True
delete.Enabled = True
savecustomer.Enabled = False
End Sub

Private Sub SSTab1_GotFocus()
ListView1.ListItems.clear
Set r1 = c1.Execute("select * from customer_detail")

If IsNull(r1.Fields(11)) Then
 rf11 = None
Else
 rf11 = r1.Fields(11)
End If

If IsNull(r1.Fields(9)) Then
 rf9 = None
Else
 rf9 = r1.Fields(9)
End If

If IsNull(r1.Fields(8)) Then
 rf8 = None
Else
 rf8 = r1.Fields(8)
End If

If IsNull(r1.Fields(10)) Then
 rf10 = None
Else
 rf10 = r1.Fields(10)
End If

While Not r1.EOF

Set item = ListView1.ListItems.add(, , r1.Fields(0))
item.SubItems(1) = r1.Fields(1)
item.SubItems(2) = r1.Fields(6)
item.SubItems(3) = r1.Fields(3)

item.SubItems(4) = r1.Fields(2)

item.SubItems(5) = rf11
item.SubItems(6) = rf10
item.SubItems(7) = rf8
item.SubItems(8) = rf9

r1.MoveNext
Wend

End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
Text4.BackColor = vbWhite
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn Then

Else
KeyAscii = 0
MsgBox "Must be a Number", vbCritical
End If
End Sub



Private Sub Text6_KeyPress(KeyAscii As Integer)
Text6.BackColor = vbWhite
End Sub

Private Sub update_Click()
ans = MsgBox("Do you Want to Update", vbOKCancel + vbInformation)
If ans = 1 Then
Set r1 = New ADODB.Recordset
sql = "update customer_detail set customer_name='" + customer_name.Text + "',gstin=" + gstin.Text + ",address='" + address.Text + "',district='" + Combo1.Text + "',state='" + Combo2.Text + "',phone_no='" + Text4.Text + "',email='" + Text6.Text + "' ,account_no='" + account_no.Text + "',holder_name='" + holder_name.Text + "',ifcs_code='" + ifsc_code.Text + "',bank_name='" + bank_name.Text + "'where customer_id='" + Combo3.Text + "'"
Set r1 = c1.Execute(sql)
MsgBox "updated"
auto
Adodc1.Refresh
End If
End Sub
Public Function additem_combo()
Set r1 = New ADODB.Recordset
sql = "select * from customer_detail where customer_id='" + Combo3.Text + "'"
Set r1 = c1.Execute(sql)
customer_id.Text = r1.Fields(0)
customer_name.Text = r1.Fields(1)
gstin.Text = r1.Fields(2)
Text4.Text = r1.Fields(6)
address.Text = r1.Fields(3)
Combo1.Text = r1.Fields(4)
Combo2.Text = r1.Fields(5)
Text6.Text = r1.Fields(7)


If IsNull(r1.Fields(11)) Then
 rf11 = None
Else
 rf11 = r1.Fields(11)
End If

If IsNull(r1.Fields(9)) Then
 rf9 = None
Else
 rf9 = r1.Fields(9)
End If

If IsNull(r1.Fields(8)) Then
 rf8 = None
Else
 rf8 = r1.Fields(8)
End If

If IsNull(r1.Fields(10)) Then
 rf10 = None
Else
 rf10 = r1.Fields(10)
End If

account_no.Text = rf8
holder_name.Text = rf9
ifsc_code.Text = rf10
bank_name.Text = rf11

End Function





