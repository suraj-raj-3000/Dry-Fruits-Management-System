VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form m_form 
   BackColor       =   &H80000004&
   Caption         =   "home"
   ClientHeight    =   9750
   ClientLeft      =   120
   ClientTop       =   810
   ClientWidth     =   16020
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9750
   ScaleWidth      =   16020
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      TabIndex        =   8
      Top             =   9065
      Width           =   20440
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   15720
         Picture         =   "m_form.frx":0000
         ScaleHeight     =   330
         ScaleWidth      =   330
         TabIndex        =   20
         Top             =   70
         Width           =   330
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   10560
         Picture         =   "m_form.frx":0530
         ScaleHeight     =   330
         ScaleWidth      =   330
         TabIndex        =   19
         Top             =   70
         Width           =   330
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   5520
         Picture         =   "m_form.frx":0A35
         ScaleHeight     =   330
         ScaleWidth      =   330
         TabIndex        =   18
         Top             =   70
         Width           =   330
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   360
         Picture         =   "m_form.frx":0F82
         ScaleHeight     =   330
         ScaleWidth      =   330
         TabIndex        =   17
         Top             =   70
         Width           =   330
      End
      Begin VB.Image Image11 
         Height          =   495
         Left            =   15360
         Top             =   0
         Width           =   5100
      End
      Begin VB.Image Image10 
         Height          =   495
         Left            =   10200
         Top             =   0
         Width           =   5100
      End
      Begin VB.Image Image9 
         Height          =   495
         Left            =   5160
         Top             =   0
         Width           =   5100
      End
      Begin VB.Image Image2 
         Height          =   495
         Left            =   0
         Top             =   0
         Width           =   5100
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
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
         Left            =   19800
         TabIndex        =   16
         Top             =   120
         Width           =   315
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
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
         Left            =   14520
         TabIndex        =   15
         Top             =   120
         Width           =   315
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
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
         Left            =   9240
         TabIndex        =   14
         Top             =   120
         Width           =   315
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
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
         Left            =   4200
         TabIndex        =   13
         Top             =   120
         Width           =   315
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Total Order  :-"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   240
         Left            =   11400
         TabIndex        =   12
         Top             =   120
         Width           =   2370
      End
      Begin VB.Line Line4 
         X1              =   20430
         X2              =   20430
         Y1              =   0
         Y2              =   600
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Pending Order   :-"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   240
         Left            =   16920
         TabIndex        =   11
         Top             =   120
         Width           =   2730
      End
      Begin VB.Line Line3 
         X1              =   15322
         X2              =   15322
         Y1              =   0
         Y2              =   600
      End
      Begin VB.Shape Shape1 
         Height          =   495
         Left            =   0
         Top             =   0
         Width           =   20430
      End
      Begin VB.Line Line2 
         X1              =   10215
         X2              =   10215
         Y1              =   0
         Y2              =   600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Pending Order   :-"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   6360
         TabIndex        =   10
         Top             =   120
         Width           =   2640
      End
      Begin VB.Line Line1 
         X1              =   5107
         X2              =   5107
         Y1              =   0
         Y2              =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Total Order   :-"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   1320
         TabIndex        =   9
         Top             =   120
         Width           =   2340
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   30
      Left            =   0
      TabIndex        =   7
      Top             =   9720
      Width           =   16020
      _ExtentX        =   28258
      _ExtentY        =   53
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Image Image7 
      Height          =   2895
      Left            =   14400
      Top             =   4920
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "RETURN"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14760
      TabIndex        =   6
      Top             =   7320
      Width           =   1935
   End
   Begin VB.Image Image6 
      Height          =   3015
      Left            =   15840
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Image Image5 
      Height          =   2895
      Left            =   9360
      Top             =   4920
      Width           =   2535
   End
   Begin VB.Image Image4 
      Height          =   3015
      Left            =   4080
      Top             =   4920
      Width           =   2655
   End
   Begin VB.Image Image3 
      Height          =   3015
      Left            =   11400
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Image Img 
      Height          =   3015
      Left            =   6840
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   3015
      Left            =   1680
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Image purchase 
      Height          =   2250
      Left            =   1800
      Picture         =   "m_form.frx":51B4
      Top             =   1560
      Width           =   2250
   End
   Begin VB.Image stock 
      Height          =   2415
      Left            =   6960
      Picture         =   "m_form.frx":BD75
      Top             =   1440
      Width           =   2340
   End
   Begin VB.Image admin_form 
      Height          =   2385
      Left            =   4200
      Picture         =   "m_form.frx":18C64
      Top             =   5160
      Width           =   2415
   End
   Begin VB.Image sell 
      Height          =   2250
      Left            =   11640
      Picture         =   "m_form.frx":20E74
      Top             =   1440
      Width           =   2250
   End
   Begin VB.Image account 
      Height          =   2250
      Left            =   15960
      Picture         =   "m_form.frx":2AF23
      Top             =   1320
      Width           =   2250
   End
   Begin VB.Image report 
      Height          =   2400
      Left            =   9480
      Picture         =   "m_form.frx":33413
      Top             =   5040
      Width           =   2400
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PURCHASE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   5
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " SELL"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11880
      TabIndex        =   4
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "STOCK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7200
      TabIndex        =   3
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " ORDER"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   16200
      TabIndex        =   2
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " REPORT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      TabIndex        =   1
      Top             =   7440
      Width           =   1935
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " ADMIN"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   0
      Top             =   7560
      Width           =   1935
   End
   Begin VB.Image Image8 
      Height          =   2250
      Left            =   14520
      Picture         =   "m_form.frx":3CB3A
      Top             =   5040
      Width           =   2250
   End
End
Attribute VB_Name = "m_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Form_Load()
Connection

 Image1.MousePointer = vbCustom
 Image1.MouseIcon = LoadPicture("C:\Program Files x86\Microsoft Visual Studio\COMMON\Graphics\Cursors\H_POINT.cur")
 
 Img.MousePointer = vbCustom
 Img.MouseIcon = LoadPicture("C:\Program Files x86\Microsoft Visual Studio\COMMON\Graphics\Cursors\H_POINT.cur")
 
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

 
 Image2.MousePointer = vbCustom
 Image2.MouseIcon = LoadPicture("C:\Program Files x86\Microsoft Visual Studio\COMMON\Graphics\Cursors\H_POINT.cur")

 Image11.MousePointer = vbCustom
 Image11.MouseIcon = LoadPicture("C:\Program Files x86\Microsoft Visual Studio\COMMON\Graphics\Cursors\H_POINT.cur")
 
 Image9.MousePointer = vbCustom
 Image9.MouseIcon = LoadPicture("C:\Program Files x86\Microsoft Visual Studio\COMMON\Graphics\Cursors\H_POINT.cur")
 
 Image10.MousePointer = vbCustom
 Image10.MouseIcon = LoadPicture("C:\Program Files x86\Microsoft Visual Studio\COMMON\Graphics\Cursors\H_POINT.cur")
 
MDIForm1.Picture2.Visible = False
m_form.Caption = "Home"

Set r = c.Execute("select count(order_number) from order_detail ")
Label12.Caption = r.Fields(0)

Set r = c.Execute("select count(order_number) from order_detail where inv_status='no'")
Label13.Caption = r.Fields(0)

Set r = c.Execute("select count(order_number) from customer_order_detail where status='no' ")
Label15.Caption = r.Fields(0)

Set r = c.Execute("select count(order_number) from customer_order_detail ")
Label14.Caption = r.Fields(0)

End Sub

Private Sub Image1_Click()
Load purchased_product
purchased_product.Show
MDIForm1.Picture2.Visible = True
End Sub

Private Sub Image10_Click()
Unload stock_form
Unload report_form
Unload product_detail
Unload report_form
Unload admin
Unload purchased_product
Unload sells
Unload sell_return
Unload purchase_return
Unload m_form
Unload supplier_order

CUSTOMER_ORDER.Show
End Sub

Private Sub Image11_Click()
Unload stock_form
Unload report_form
Unload product_detail
Unload report_form
Unload admin
Unload sell_return
Unload purchase_return
Unload m_form
Unload supplier_order
Unload CUSTOMER_ORDER
Unload purchased_product

sells.Show
End Sub

Private Sub Image2_Click()
Unload CUSTOMER_ORDER
Unload stock_form
Unload report_form
Unload product_detail
Unload report_form
Unload admin
Unload purchased_product
Unload sells
Unload sell_return
Unload purchase_return
Unload m_form

supplier_order.Show
supplier_order.Frame1.Visible = True
supplier_order.Frame2.Visible = False
supplier_order.Frame3.Visible = False
End Sub

Private Sub Image3_Click()
sells.Show
MDIForm1.Picture2.Visible = True
End Sub

Private Sub Image4_Click()
admin.Show
End Sub

Private Sub Image5_Click()
MDIForm1.Picture2.Visible = True
report_form.Show
End Sub

Private Sub Image6_Click()
order.Show
End Sub

Private Sub Image7_Click()
MDIForm1.Picture2.Visible = True
return_form.Show
End Sub

Private Sub Image9_Click()
Unload stock_form
Unload report_form
Unload product_detail
Unload report_form
Unload admin
Unload sells
Unload sell_return
Unload purchase_return
Unload m_form
Unload supplier_order
Unload CUSTOMER_ORDER

purchased_product.Show
End Sub

Private Sub Img_Click()
stock_form.Show
End Sub

