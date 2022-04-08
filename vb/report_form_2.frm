VERSION 5.00
Begin VB.Form report_form 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   9615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17175
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9615
   ScaleWidth      =   17175
   WindowState     =   2  'Maximized
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "REPORT"
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
      TabIndex        =   8
      Top             =   0
      Width           =   9045
   End
   Begin VB.Image Image16 
      Height          =   2895
      Left            =   13200
      Top             =   5880
      Width           =   3375
   End
   Begin VB.Image Image15 
      Height          =   2895
      Left            =   9240
      Top             =   5880
      Width           =   3615
   End
   Begin VB.Image Image14 
      Height          =   2895
      Left            =   5040
      Top             =   5880
      Width           =   3495
   End
   Begin VB.Image Image13 
      Height          =   2895
      Left            =   840
      Top             =   5880
      Width           =   3255
   End
   Begin VB.Image Image12 
      Height          =   2895
      Left            =   13440
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Image Image11 
      Height          =   2895
      Left            =   9480
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Image Image10 
      Height          =   2895
      Left            =   5160
      Top             =   1560
      Width           =   3615
   End
   Begin VB.Image Image9 
      Height          =   2895
      Left            =   1200
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " CUSTOMER REPORT"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9090
      TabIndex        =   7
      Top             =   8400
      Width           =   3765
   End
   Begin VB.Image Image8 
      Height          =   2235
      Left            =   9960
      Picture         =   "report_form.frx":0000
      Top             =   5880
      Width           =   2250
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " SUPPLIER REPORT"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5025
      TabIndex        =   6
      Top             =   8400
      Width           =   3465
   End
   Begin VB.Image Image7 
      Height          =   2250
      Left            =   5760
      Picture         =   "report_form.frx":2700
      Top             =   6000
      Width           =   2250
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCT REPORT"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   13215
      TabIndex        =   5
      Top             =   8400
      Width           =   3345
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "STOCK REPORT"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9510
      TabIndex        =   4
      Top             =   4080
      Width           =   2835
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RETURN REPORT"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   975
      TabIndex        =   3
      Top             =   8400
      Width           =   3105
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SELL REPORT"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   13740
      TabIndex        =   2
      Top             =   4080
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " PURCHASE REPORT"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5130
      TabIndex        =   1
      Top             =   4080
      Width           =   3675
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ORDER REPORT"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1200
      TabIndex        =   0
      Top             =   4080
      Width           =   2895
   End
   Begin VB.Image Image6 
      Height          =   2250
      Left            =   13920
      Picture         =   "report_form.frx":7F8B
      Top             =   6000
      Width           =   2250
   End
   Begin VB.Image Image5 
      Height          =   2250
      Left            =   9840
      Picture         =   "report_form.frx":D634
      Top             =   1680
      Width           =   2250
   End
   Begin VB.Image Image4 
      Height          =   2250
      Left            =   13800
      Picture         =   "report_form.frx":FCC0
      Top             =   1680
      Width           =   2250
   End
   Begin VB.Image Image3 
      Height          =   2250
      Left            =   1320
      Picture         =   "report_form.frx":12549
      Top             =   6000
      Width           =   2250
   End
   Begin VB.Image Image2 
      Height          =   2340
      Left            =   5880
      Picture         =   "report_form.frx":14FFD
      Top             =   1560
      Width           =   2250
   End
   Begin VB.Image Image1 
      Height          =   2250
      Left            =   1560
      Picture         =   "report_form.frx":1756F
      Top             =   1680
      Width           =   2250
   End
End
Attribute VB_Name = "report_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Form_Load()
report_form.Caption = "Report"
End Sub

Private Sub Image10_Click()
Unload report
purchase_report_function
End Sub

Private Sub Image11_Click()
Unload report
stock_report_function
End Sub

Private Sub Image12_Click()
Unload report
sale_report_function
End Sub

Private Sub Image13_Click()
Unload report
return_report_function
End Sub

Private Sub Image14_Click()
Unload report
supplier_report_function
End Sub

Private Sub Image15_Click()
Unload report
customer_report_function
End Sub

Private Sub Image16_Click()
Unload report
product_report_function
End Sub

Private Sub Image9_Click()
Unload report
order_report_function
End Sub
