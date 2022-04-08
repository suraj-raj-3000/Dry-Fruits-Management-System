VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form product_detail 
   BackColor       =   &H80000004&
   Caption         =   "Form1"
   ClientHeight    =   9525
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13170
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9525
   ScaleWidth      =   13170
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   7095
      Left            =   3000
      TabIndex        =   13
      Top             =   840
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   12515
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "PRODUCT"
      TabPicture(0)   =   "product_detail.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "VIEW"
      TabPicture(1)   =   "product_detail.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ListView1"
      Tab(1).Control(1)=   "Label26"
      Tab(1).Control(2)=   "Label25"
      Tab(1).Control(3)=   "Label10"
      Tab(1).ControlCount=   4
      Begin MSComctlLib.ListView ListView1 
         Height          =   5055
         Left            =   -74400
         TabIndex        =   29
         Top             =   1080
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   8916
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
            Text            =   "S.N"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "NAME"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "BRAND"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "UNIT"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "GST"
            Object.Width           =   3175
         EndProperty
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Caption         =   " Product Information"
         Height          =   5535
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   10935
         Begin VB.ListBox List1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   2760
            ItemData        =   "product_detail.frx":0038
            Left            =   6240
            List            =   "product_detail.frx":003A
            TabIndex        =   17
            Top             =   2040
            Width           =   1575
         End
         Begin VB.CommandButton delete_brand 
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
            Height          =   495
            Left            =   9600
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   3120
            Width           =   1095
         End
         Begin VB.TextBox gst 
            Height          =   400
            Left            =   2280
            MaxLength       =   3
            TabIndex        =   1
            Top             =   2640
            Width           =   2175
         End
         Begin VB.TextBox text1 
            Height          =   400
            Left            =   2280
            MaxLength       =   30
            TabIndex        =   0
            Top             =   1920
            Width           =   2175
         End
         Begin VB.TextBox Text3 
            Height          =   400
            Left            =   2280
            MaxLength       =   15
            TabIndex        =   2
            Top             =   3360
            Width           =   2175
         End
         Begin VB.CommandButton Command1 
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
            Height          =   495
            Left            =   9600
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   2160
            Width           =   1095
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
            ItemData        =   "product_detail.frx":003C
            Left            =   2280
            List            =   "product_detail.frx":004C
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   4080
            Width           =   2175
         End
         Begin VB.ListBox List4 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   2760
            ItemData        =   "product_detail.frx":0068
            Left            =   7800
            List            =   "product_detail.frx":006A
            TabIndex        =   16
            Top             =   2040
            Width           =   1335
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
            ItemData        =   "product_detail.frx":006C
            Left            =   2280
            List            =   "product_detail.frx":007C
            TabIndex        =   5
            Top             =   4800
            Width           =   2175
         End
         Begin VB.TextBox product_id 
            DataSource      =   "Adodc1"
            Enabled         =   0   'False
            Height          =   400
            Left            =   2280
            TabIndex        =   18
            Text            =   " "
            Top             =   1200
            Width           =   2175
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
            Left            =   2280
            TabIndex        =   15
            Text            =   "Select Product ID"
            Top             =   1200
            Visible         =   0   'False
            Width           =   2190
         End
         Begin VB.ListBox List3 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   2760
            ItemData        =   "product_detail.frx":0098
            Left            =   5160
            List            =   "product_detail.frx":009A
            TabIndex        =   19
            Top             =   2040
            Width           =   1095
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
            Left            =   1920
            TabIndex        =   39
            Top             =   4920
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
            Left            =   1440
            TabIndex        =   38
            Top             =   4200
            Width           =   120
         End
         Begin VB.Label Label13 
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
            TabIndex        =   37
            Top             =   3480
            Width           =   120
         End
         Begin VB.Label Label12 
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
            TabIndex        =   36
            Top             =   2760
            Width           =   120
         End
         Begin VB.Label Label11 
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
            TabIndex        =   35
            Top             =   2040
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
            Left            =   1800
            TabIndex        =   34
            Top             =   1200
            Width           =   120
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
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
            Left            =   240
            TabIndex        =   28
            Top             =   2040
            Width           =   1545
         End
         Begin VB.Label Label3 
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
            Left            =   840
            TabIndex        =   27
            Top             =   3480
            Width           =   690
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Product ID"
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
            TabIndex        =   26
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " Gst %"
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
            TabIndex        =   25
            Top             =   2760
            Width           =   675
         End
         Begin VB.Shape Shape2 
            Height          =   255
            Left            =   5160
            Top             =   1800
            Width           =   3975
         End
         Begin VB.Line Line5 
            X1              =   6240
            X2              =   6240
            Y1              =   2040
            Y2              =   1800
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
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
            Left            =   5520
            TabIndex        =   24
            Top             =   1815
            Width           =   405
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
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
            Left            =   6750
            TabIndex        =   23
            Top             =   1815
            Width           =   585
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   840
            TabIndex        =   22
            Top             =   4200
            Width           =   480
         End
         Begin VB.Line Line6 
            X1              =   7800
            X2              =   7800
            Y1              =   2040
            Y2              =   1800
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   8265
            TabIndex        =   21
            Top             =   1815
            Width           =   375
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " Unit Values  "
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
            TabIndex        =   20
            Top             =   4920
            Width           =   1380
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
         Left            =   -67920
         TabIndex        =   33
         Top             =   720
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
         Left            =   -70680
         TabIndex        =   32
         Top             =   720
         Width           =   165
      End
      Begin VB.Label Label10 
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
         Left            =   -70440
         TabIndex        =   31
         Top             =   720
         Width           =   2370
      End
   End
   Begin VB.CommandButton Command2 
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
      Left            =   3360
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
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8640
      Width           =   1695
   End
   Begin VB.CommandButton saveproduct 
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
      TabIndex        =   8
      Top             =   8640
      Width           =   1695
   End
   Begin VB.CommandButton viewproduct 
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
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   11
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
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8640
      Width           =   1695
   End
   Begin VB.TextBox stock 
      Height          =   285
      Left            =   3000
      TabIndex        =   12
      Text            =   "Text2"
      Top             =   5760
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "PRODUCT  DETAIL"
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
      Left            =   4080
      TabIndex        =   30
      Top             =   120
      Width           =   9045
   End
End
Attribute VB_Name = "product_detail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pk As String
Dim c As ADODB.Connection
Dim r As ADODB.Recordset
Dim sql As String
Dim i As Integer
Dim item As ListItem
Public ind As Integer




Private Sub Combo2_Click()
Combo3.clear
If Combo2.Text = "packets" Then
Combo3.AddItem "250 gram"
Combo3.AddItem "500 gram"
Combo3.AddItem "1 kg"
Combo3.AddItem "2.5 kg"
ElseIf Combo2.Text = "Bags" Then
Combo3.AddItem "5 kg"
Combo3.AddItem "10 kg"
Combo3.AddItem "25 kg"
Combo3.AddItem "50 kg"
ElseIf Combo2.Text = "Box" Then
Combo3.AddItem "25 pieces/4 kg"
Combo3.AddItem "50 pieces/3 kg"
Combo3.AddItem "100 pieces /2 kg"
Combo3.AddItem "200 pieces /1 kg"
ElseIf Combo2.Text = "Kg" Then
Combo3.AddItem "1"
Combo3.AddItem "2"
Combo3.AddItem "5"
Combo3.AddItem "10"
Combo3.AddItem "20"
Combo3.AddItem "50"
End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
Combo2.BackColor = vbWhite
End Sub


Private Sub Combo3_KeyPress(KeyAscii As Integer)
Combo3.BackColor = vbWhite
End Sub

Private Sub Command1_Click()
If Text3.Text = "" Or Combo2.Text = "" Or Combo2.Text = "combo2" Or Combo3.Text = "" Then
 MsgBox "Select Brand and Unit", vbCritical, "Warning"
Else
List1.BackColor = vbWhite
List4.BackColor = vbWhite
List1.AddItem Text3.Text
Set r = c.Execute("select max(s_no) from product_brand")
snn = r.Fields(0) + 1
List3.AddItem snn
List4.AddItem Combo3.Text + " " + Combo2.Text
End If
i = i + 1

End Sub


Private Sub brand_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Combo2.SetFocus
End If
End Sub

Private Sub Combo1_Click()
additem_combo
End Sub

Private Sub Command2_Click()
autogenerate
gst.Text = ""
Text1.Text = ""
Text3.Text = ""
List1.clear
List3.clear
List4.clear
Combo1.Visible = False
product_id.Visible = True
saveproduct.Enabled = True

update.Enabled = False
saveproduct.Enabled = True
delete.Enabled = False

End Sub



Private Sub delete_brand_Click()
If List1.List(ind) = "" Then
Else
 List1.RemoveItem (ind)
 List3.RemoveItem (ind)
 List4.RemoveItem (ind)
End If
End Sub
Private Sub Command3_Click()
List1.RemoveItem i
List1.AddItem Text3.Text
End Sub

Private Sub delete_Click()
Dim i As Integer
ans = MsgBox("Do you Want to Delete", vbOKCancel + vbInformation)
If ans = 1 Then
sql = "delete from product_brand where product_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)

sql = "delete from product_detail where product_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)
MsgBox "Product deleted"
List1.clear
List2.clear
Text1.Text = ""
Combo2.Text = ""
gst.Text = ""
List1.clear
List2.clear
List3.clear
List4.clear
End If
End Sub




Private Sub Form_Load()

Connection
autogenerate

product_detail.Caption = "Product Entry "
all_product
End Sub



Private Sub gst_KeyPress(KeyAscii As Integer)
gst.BackColor = vbWhite
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn Then
  If KeyAscii = 13 Then
   saveproduct.SetFocus
  End If
Else
KeyAscii = 0
MsgBox "Enter only number"
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
Unload product_detail
Unload supplier_entry_form

Load customer_entry_form
customer_entry_form.Show
End Sub



Private Sub price_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn Then
  If KeyAscii = 13 Then
  gst.SetFocus
  End If
Else
KeyAscii = 0
MsgBox "Enter only number"
End If
End Sub









Private Sub product_name_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
brand.SetFocus
End If
End Sub

Private Sub List1_Click()
Dim i As Integer
i = List1.ListIndex
Text3.Text = List1.Text

ind = List1.ListIndex

End Sub

Private Sub OK_Click()
Dim i As Integer
i = List1.ListIndex
List1.List(i) = Text2.Text

End Sub





Private Sub List3_Click()
i = List1.ListIndex
Text3.Text = List1.List(i)
ind = List3.ListIndex

End Sub

Private Sub List4_Click()
ind = List4.ListIndex
End Sub
Private Sub saveproduct_Click()
Dim i As Integer
ans = MsgBox("Do you Want to Save", vbOKCancel + vbInformation)
If ans = 1 Then

If Text1.Text = "" Then
 Text1.BackColor = &HC0C0FF
 MsgBox "Product Name Fields is Empty", vbCritical
ElseIf gst.Text = "" Then
 gst.BackColor = &HC0C0FF
 MsgBox "Gst Fields is Empty", vbCritical
ElseIf List1.List(0) = "" Then
List1.BackColor = &HC0C0FF
MsgBox "Add Brand and unit in Listbox", vbCritical
ElseIf List3.List(0) = "" Then
 List3.BackColor = &HC0C0FF
 MsgBox "Add Brand and unit in Listbox", vbCritical
ElseIf List4.List(0) = "" Then
 List4.BackColor = &HC0C0FF
 MsgBox "Add Brand and unit in Listbox", vbCritical
Else
sql = "insert into product_detail values('" + product_id.Text + "','" + Text1.Text + "'," + gst.Text + ",NULL,NULL)"

Set r = c.Execute(sql)


MsgBox "Product saved"


For i = 0 To List1.ListCount - 1
autostock
sql = "insert into stock_detail values('" + stock.Text + "',NULL,'" + product_id.Text + "','" + Text1.Text + "','" + List4.List(i) + "','" + "0" + "',NULL,'" + List1.List(i) + "')"
Set r = c.Execute(sql)
Next i


Text1.Text = ""
Text3.Text = ""
Combo3.Text = ""
gst.Text = ""
autogenerate
List1.clear
List3.clear
List4.clear
End If
End If
ListView1.ListItems.clear
all_product
End Sub





Public Function additem_combo()
Set r = New ADODB.Recordset
sql = "select * from product_detail where product_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)

product_id.Text = r.Fields(0)
Text1.Text = r.Fields(1)
gst.Text = r.Fields(2)

List1.clear
List2.clear
List3.clear
List4.clear

Set r = New ADODB.Recordset
sql = "select  * from product_brand where product_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)
Text3.Text = r.Fields(0)
Do While Not r.EOF
List1.AddItem r!brand
List2.AddItem r!s_no
List3.AddItem i + 1
List4.AddItem r.Fields(3)
r.MoveNext
i = i + 1
Loop

End Function

Public Function Connection()
Set c = New ADODB.Connection
c.Open "Provider=MSDAORA.1;Password=lnt123;User ID=LNT;Persist Security Info=True"
Set r = New ADODB.Recordset
End Function

Public Function autogenerate()
Dim a As String
product_id.Visible = True

Set r = New ADODB.Recordset
sql = "select max(to_number(substr(product_id,5,length(product_id)))) from product_detail"
Set r = c.Execute(sql)
If IsNull(r.Fields(0)) Then
product_id.Text = "PI" & "00" & 1
Else
product_id.Text = "PI" & "00" & r.Fields(0) + 1
End If
a = product_id.Text
If (a = "PI" & "001" & "0") Then
sql = "select max(to_number(substr(product_id,4,length(product_id)))) from product_detail"
Set r = c.Execute(sql)
product_id.Text = "PI" & "0" & r.Fields(0) + 1
End If

End Function






Private Sub Text1_KeyPress(KeyAscii As Integer)
Text1.BackColor = vbWhite
Select Case KeyAscii
 Case 32 To 64, 91 To 96, 123 To 126
  MsgBox "Must be a latter", vbCritical, "Warning"
  KeyAscii = 0
End Select
End Sub

Private Sub text1_LostFocus()
Text1.Text = UCase(Text1.Text)
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
Text3.BackColor = vbWhite
Select Case KeyAscii
 Case 33 To 64, 91 To 96, 123 To 126
  MsgBox "Must be a latter", vbCritical, "Warning"
  KeyAscii = 0
End Select
End Sub

Private Sub Text3_lostfocus()
Text3.Text = UCase(Text3.Text)
End Sub

Private Sub update_Click()
Dim i As Integer
ans = MsgBox("Do you Want to Update", vbOKCancel + vbInformation)
If ans = 1 Then
For i = 0 To List3.ListCount - 1
Set r = New ADODB.Recordset
sql = "update product_brand set brand='" + List1.List(i) + "',unit='" + List4.List(i) + "' where s_no='" + List2.List(i) + "'"
Set r = c.Execute(sql)
Next i

Set r = New ADODB.Recordset
sql = "update product_detail set product_name='" + Text1.Text + "',gst='" + gst.Text + "' where product_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)

MsgBox "Product updated"
End If
additem_combo
End Sub



Private Sub viewproduct_Click()
update.Enabled = True
delete.Enabled = True
product_id.Visible = False
Combo1.Visible = True

Combo1.clear
sql = "select product_id from product_detail"
Set r = c.Execute(sql)
While r.EOF = False
Combo1.AddItem r.Fields(0)
r.MoveNext
Wend
saveproduct.Enabled = False
update.Enabled = True
delete.Enabled = True
End Sub

Public Function autostock()

Dim a As String
Set r = New ADODB.Recordset
sql = "select max(to_number(substr(stock_no,5,length(stock_no)))) from stock_detail"
Set r = c.Execute(sql)
If IsNull(r.Fields(0)) Then
stock.Text = "st" & "00" & 1
Else
stock.Text = "st" & "00" & r.Fields(0) + 1
End If
b = stock.Text
If (a = "st" & "001" & "0") Then
sql = "select max(to_number(substr(stock_no,4,length(stock_no)))) from stock_detail"
Set r = c.Execute(sql)
stock.Text = "st" & "0" & r.Fields(0) + 1
End If
End Function

Public Function brand_no()
j = List2.ListCount
Set r = New ADODB.Recordset
sql = "select max(to_number(substr(s_no,5,length(s_no)))) from brandsno"
Set r = c.Execute(sql)
If IsNull(r.Fields(0)) Then
List2.AddItem "sn" & "00" & 1
Else
List2.AddItem "sn" & "0" & r.Fields(0) + 1
End If
a = List2.List(j)
If (a = "sn" & "01" & "0") Then
sql = "select max(to_number(substr(s_no,4,length(s_no)))) from brandsno"
Set r = c.Execute(sql)
List2.List(j) = "sn" & "0" & r.Fields(0) + 1
End If
End Function

Public Sub all_product()
ListView1.ListItems.clear
Set r = c.Execute("select * from product_detail,product_brand")
While Not r.EOF

Set item = ListView1.ListItems.add(, , r.Fields(0))
item.SubItems(1) = r.Fields(0)
item.SubItems(2) = r.Fields(1)
item.SubItems(3) = r.Fields(5)

item.SubItems(4) = r.Fields(8)
item.SubItems(5) = r.Fields(2)
'item.SubItems(6) = r.Fields(0)
 
r.MoveNext
Wend
End Sub
