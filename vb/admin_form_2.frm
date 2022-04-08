VERSION 5.00
Begin VB.Form admin 
   Caption         =   "Admin"
   ClientHeight    =   8460
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16050
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8460
   ScaleWidth      =   16050
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   10935
      Left            =   0
      Picture         =   "admin_form.frx":0000
      ScaleHeight     =   10875
      ScaleWidth      =   18435
      TabIndex        =   0
      Top             =   0
      Width           =   18495
      Begin VB.Image Image4 
         Height          =   3855
         Left            =   9240
         Top             =   5160
         Width           =   6015
      End
      Begin VB.Image Image3 
         Height          =   3855
         Left            =   2520
         Top             =   5160
         Width           =   6015
      End
      Begin VB.Image Image2 
         Height          =   3855
         Left            =   9240
         Top             =   600
         Width           =   6015
      End
      Begin VB.Image Image1 
         Height          =   3855
         Left            =   2520
         Top             =   600
         Width           =   6015
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   30
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   765
         Left            =   2520
         TabIndex        =   4
         Top             =   3360
         Width           =   5850
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   30
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   645
         Left            =   9720
         TabIndex        =   3
         Top             =   3480
         Width           =   5250
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "Rreport"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   30
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   690
         Left            =   2880
         TabIndex        =   2
         Top             =   8040
         Width           =   5520
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   30
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   690
         Left            =   9480
         TabIndex        =   1
         Top             =   8160
         Width           =   5520
      End
   End
End
Attribute VB_Name = "admin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Form_Load()
MDIForm1.Picture2.Visible = True
Image1.MousePointer = vbCustom
Image1.MouseIcon = LoadPicture("C:\Program Files x86\Microsoft Visual Studio\COMMON\Graphics\Cursors\H_POINT.cur")

Image4.MousePointer = vbCustom
Image4.MouseIcon = LoadPicture("C:\Program Files x86\Microsoft Visual Studio\COMMON\Graphics\Cursors\H_POINT.cur")

Image2.MousePointer = vbCustom
Image2.MouseIcon = LoadPicture("C:\Program Files x86\Microsoft Visual Studio\COMMON\Graphics\Cursors\H_POINT.cur")

Image3.MousePointer = vbCustom
Image3.MouseIcon = LoadPicture("C:\Program Files x86\Microsoft Visual Studio\COMMON\Graphics\Cursors\H_POINT.cur")

End Sub



Private Sub Image1_Click()
Load supplier_entry_form
supplier_entry_form.Show
End Sub

Private Sub Image2_Click()
Load product_detail
product_detail.Show
End Sub

Private Sub Image3_Click()
report.Show
End Sub

Private Sub Image4_Click()
Load customer_entry_form
customer_entry_form.Show
End Sub

Private Sub Label3_Click()
Load customer_entry_form
customer_entry_form.Show
End Sub


