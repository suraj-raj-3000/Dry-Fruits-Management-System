VERSION 5.00
Begin VB.Form return_form 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   6960
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12720
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.Image Image4 
      Height          =   3975
      Left            =   11040
      Top             =   2040
      Width           =   3615
   End
   Begin VB.Image Image3 
      Height          =   3975
      Left            =   4560
      Top             =   1920
      Width           =   3615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      Caption         =   "Sell Return"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   525
      Left            =   11775
      TabIndex        =   1
      Top             =   5400
      Width           =   2310
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      Caption         =   "Purchase Return"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   525
      Left            =   4635
      TabIndex        =   0
      Top             =   5280
      Width           =   3390
   End
   Begin VB.Image Image2 
      Height          =   3000
      Left            =   11400
      Picture         =   "return_form.frx":0000
      Top             =   2160
      Width           =   3000
   End
   Begin VB.Image Image1 
      Height          =   3000
      Left            =   4920
      Picture         =   "return_form.frx":8226
      Top             =   2040
      Width           =   3000
   End
End
Attribute VB_Name = "return_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Image3.MousePointer = vbCustom
 Image3.MouseIcon = LoadPicture("C:\Program Files x86\Microsoft Visual Studio\COMMON\Graphics\Cursors\H_POINT.cur")
 
 Image4.MousePointer = vbCustom
 Image4.MouseIcon = LoadPicture("C:\Program Files x86\Microsoft Visual Studio\COMMON\Graphics\Cursors\H_POINT.cur")
 
 return_form.Caption = "Return"
End Sub

Private Sub Image3_Click()
purchase_return.Show
End Sub

Private Sub Image4_Click()
sell_return.Show
End Sub
