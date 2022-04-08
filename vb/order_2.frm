VERSION 5.00
Begin VB.Form order 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   7470
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14580
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7470
   ScaleWidth      =   14580
   WindowState     =   2  'Maximized
   Begin VB.Image Image4 
      Height          =   3975
      Left            =   11160
      Top             =   2040
      Width           =   4215
   End
   Begin VB.Image Image3 
      Height          =   3975
      Left            =   3720
      Top             =   2160
      Width           =   3975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      Caption         =   "Customer Order"
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
      Left            =   11220
      TabIndex        =   1
      Top             =   5400
      Width           =   4140
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      Caption         =   "Supplier Order"
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
      Left            =   3855
      TabIndex        =   0
      Top             =   5400
      Width           =   3780
   End
   Begin VB.Image Image2 
      Height          =   3000
      Left            =   11880
      Picture         =   "order.frx":0000
      Top             =   2280
      Width           =   3000
   End
   Begin VB.Image Image1 
      Height          =   3000
      Left            =   4200
      Picture         =   "order.frx":3166
      Top             =   2280
      Width           =   3000
   End
End
Attribute VB_Name = "order"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
Image4.MousePointer = vbCustom
 Image4.MouseIcon = LoadPicture("C:\Program Files x86\Microsoft Visual Studio\COMMON\Graphics\Cursors\H_POINT.cur")
 
 Image3.MousePointer = vbCustom
 Image3.MouseIcon = LoadPicture("C:\Program Files x86\Microsoft Visual Studio\COMMON\Graphics\Cursors\H_POINT.cur")
 
 MDIForm1.Picture2.Visible = True
 
 order.Caption = "Order"
End Sub

Private Sub Image3_Click()
Unload CUSTOMER_ORDER
Unload stock_form
Unload report_form
Unload product_detail
Unload report_form
Unload admin
Unload purchased_product
Unload sells
Unload report
Unload sell_return
Unload purchase_return
supplier_order.Show
End Sub

Private Sub Image4_Click()
CUSTOMER_ORDER.Show
End Sub
