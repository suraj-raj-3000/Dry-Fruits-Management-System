VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   ClientHeight    =   7755
   ClientLeft      =   2025
   ClientTop       =   2595
   ClientWidth     =   11280
   Icon            =   "mdiform1.frx":0000
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   4  'Align Right
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   7095
      Left            =   9285
      ScaleHeight     =   7095
      ScaleWidth      =   1995
      TabIndex        =   13
      Top             =   660
      Visible         =   0   'False
      Width           =   1995
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   720
         Top             =   2760
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Height          =   2535
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   1935
         Begin VB.Line Line1 
            BorderColor     =   &H000000C0&
            BorderStyle     =   3  'Dot
            BorderWidth     =   2
            X1              =   580
            X2              =   1455
            Y1              =   2250
            Y2              =   2250
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "LOG OUT"
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   210
            Left            =   600
            TabIndex        =   17
            Top             =   2040
            Width           =   855
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "suraj150399@gmail.com"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Suraj Kumar"
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
            TabIndex        =   15
            Top             =   240
            Width           =   1260
         End
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   4920
      Top             =   3120
   End
   Begin VB.PictureBox Picture2 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   7095
      Left            =   0
      ScaleHeight     =   7095
      ScaleWidth      =   2535
      TabIndex        =   0
      Top             =   660
      Width           =   2530
      Begin VB.Image Image17 
         Height          =   615
         Left            =   0
         Top             =   8160
         Width           =   2535
      End
      Begin VB.Image Image16 
         Height          =   615
         Left            =   0
         Top             =   7080
         Width           =   2535
      End
      Begin VB.Image Image15 
         Height          =   615
         Left            =   0
         Top             =   5880
         Width           =   2535
      End
      Begin VB.Image Image14 
         Height          =   615
         Left            =   0
         Top             =   4800
         Width           =   2535
      End
      Begin VB.Image Image13 
         Height          =   615
         Left            =   0
         Top             =   3720
         Width           =   2535
      End
      Begin VB.Image Image12 
         Height          =   615
         Left            =   0
         Top             =   2640
         Width           =   2535
      End
      Begin VB.Image Image11 
         Height          =   615
         Left            =   0
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Image Image10 
         Height          =   615
         Left            =   0
         Top             =   480
         Width           =   2535
      End
      Begin VB.Image Image8 
         Height          =   375
         Left            =   360
         Picture         =   "mdiform1.frx":EFFA
         Top             =   7200
         Width           =   375
      End
      Begin VB.Image Image3 
         Height          =   525
         Left            =   240
         Picture         =   "mdiform1.frx":F56F
         Top             =   3720
         Width           =   525
      End
      Begin VB.Label return_button 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  RETURN"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   315
         Index           =   2
         Left            =   795
         TabIndex        =   9
         Top             =   7200
         Width           =   1425
      End
      Begin VB.Shape Shape8 
         Height          =   615
         Left            =   0
         Top             =   7080
         Width           =   2535
      End
      Begin VB.Label order_button 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ORDER"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   315
         Index           =   1
         Left            =   975
         TabIndex        =   8
         Top             =   2760
         Width           =   1065
      End
      Begin VB.Image Image7 
         Height          =   375
         Left            =   240
         Picture         =   "mdiform1.frx":FBB9
         Top             =   2760
         Width           =   375
      End
      Begin VB.Shape Shape7 
         Height          =   615
         Left            =   0
         Top             =   2640
         Width           =   2535
      End
      Begin VB.Shape Shape6 
         Height          =   615
         Left            =   0
         Top             =   8160
         Width           =   2535
      End
      Begin VB.Image Image6 
         Height          =   525
         Left            =   240
         Picture         =   "mdiform1.frx":10115
         Top             =   8160
         Width           =   525
      End
      Begin VB.Label report_button 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   " REPORT"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   375
         Index           =   0
         Left            =   840
         TabIndex        =   7
         Top             =   8280
         Width           =   1455
      End
      Begin VB.Shape Shape5 
         Height          =   615
         Left            =   0
         Top             =   5880
         Width           =   2535
      End
      Begin VB.Shape Shape4 
         Height          =   615
         Left            =   0
         Top             =   4800
         Width           =   2535
      End
      Begin VB.Shape Shape3 
         Height          =   615
         Left            =   0
         Top             =   3720
         Width           =   2535
      End
      Begin VB.Shape Shape2 
         Height          =   615
         Left            =   0
         Top             =   480
         Width           =   2535
      End
      Begin VB.Image Image5 
         Height          =   450
         Left            =   240
         Picture         =   "mdiform1.frx":10696
         Top             =   6000
         Width           =   450
      End
      Begin VB.Label sell_button 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   " SELL"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   375
         Index           =   0
         Left            =   840
         TabIndex        =   6
         Top             =   6000
         Width           =   1215
      End
      Begin VB.Image Image4 
         Height          =   525
         Left            =   240
         Picture         =   "mdiform1.frx":14297
         Top             =   4800
         Width           =   525
      End
      Begin VB.Label stock_button 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   " STOCK"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   375
         Index           =   0
         Left            =   840
         TabIndex        =   5
         Top             =   4920
         Width           =   1335
      End
      Begin VB.Label purchase_button 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BUY"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   315
         Index           =   0
         Left            =   1170
         TabIndex        =   4
         Top             =   3840
         Width           =   675
      End
      Begin VB.Image Image1 
         Height          =   525
         Left            =   240
         Picture         =   "mdiform1.frx":1485B
         Top             =   480
         Width           =   525
      End
      Begin VB.Label home_button 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "HOME"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   375
         Index           =   1
         Left            =   1080
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
      Begin VB.Shape Shape1 
         Height          =   615
         Left            =   0
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Image Image2 
         Height          =   525
         Left            =   240
         Picture         =   "mdiform1.frx":172A4
         Top             =   1560
         Width           =   525
      End
      Begin VB.Label admin_button 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ADMIN"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   375
         Index           =   1
         Left            =   960
         TabIndex        =   2
         Top             =   1680
         Width           =   1215
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1800
      Top             =   1800
   End
   Begin VB.PictureBox Toolbar1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      Height          =   660
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   11220
      TabIndex        =   1
      Top             =   0
      Width           =   11280
      Begin VB.Image Image19 
         Height          =   600
         Left            =   120
         Picture         =   "mdiform1.frx":1784D
         Top             =   0
         Width           =   600
      End
      Begin VB.Image Image18 
         Height          =   810
         Left            =   18960
         Picture         =   "mdiform1.frx":1B72C
         Top             =   0
         Width           =   885
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   15600
         TabIndex        =   12
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   17040
         TabIndex        =   11
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LAXMI NARAYAN TRADERS"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   660
         Left            =   7320
         TabIndex        =   10
         Top             =   0
         Width           =   6855
      End
      Begin VB.Image Image9 
         Height          =   600
         Left            =   18240
         Picture         =   "mdiform1.frx":1BF6A
         Top             =   0
         Width           =   600
      End
   End
   Begin VB.Menu home_ 
      Caption         =   "Home"
   End
   Begin VB.Menu admin_ 
      Caption         =   "Admin"
      Begin VB.Menu supplier_entry 
         Caption         =   "Supplier Entry"
      End
      Begin VB.Menu customer_entry 
         Caption         =   "Customer Entry"
      End
      Begin VB.Menu product_entry 
         Caption         =   "Product Entry"
      End
   End
   Begin VB.Menu order_ 
      Caption         =   "Order"
      Begin VB.Menu supplier_order_ 
         Caption         =   "Supplier Order"
      End
      Begin VB.Menu customer_order_ 
         Caption         =   "Customer Order"
      End
   End
   Begin VB.Menu purchase_ 
      Caption         =   "Purchase"
   End
   Begin VB.Menu stock_ 
      Caption         =   "Stock"
   End
   Begin VB.Menu sell_ 
      Caption         =   "Sell"
   End
   Begin VB.Menu return_ 
      Caption         =   "Return"
      Begin VB.Menu purchase_return_ 
         Caption         =   "Purchase Return"
      End
      Begin VB.Menu sell_return_ 
         Caption         =   "Sell Return"
      End
   End
   Begin VB.Menu report_ 
      Caption         =   "Report"
      Begin VB.Menu order_report 
         Caption         =   "Order Report"
      End
      Begin VB.Menu purchase_report 
         Caption         =   "Purchase Report"
      End
      Begin VB.Menu sell_report 
         Caption         =   "Sale Report"
      End
      Begin VB.Menu stock_report 
         Caption         =   "Stock Report"
      End
      Begin VB.Menu product_report 
         Caption         =   "Product Report "
      End
      Begin VB.Menu customer_report 
         Caption         =   "Customer Report"
      End
      Begin VB.Menu supplier_report 
         Caption         =   "Supplier Report"
      End
      Begin VB.Menu return_report 
         Caption         =   "Return Report"
      End
   End
   Begin VB.Menu help_ 
      Caption         =   "Help"
      Begin VB.Menu about_ 
         Caption         =   "About"
      End
      Begin VB.Menu ts 
         Caption         =   "Contact Us"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub Browse_Folder()
    Dim stempDir As String
    On Error Resume Next
    stempDir = CurDir 'Current Directory
    ccc.DialogTitle = "Select A Folder "   ''' ccc is common dialog box.
    ccc.FileName = "Select Folder"
    ccc.Flags = cdlOFNNoValidate + cdlOFNHideReadOnly
    ccc.Filter = "Directories|*.~#~"
    ccc.CancelError = True
    ccc.ShowOpen
    If Err <> 32755 Then
        Label8.Caption = CurDir
    End If
    ChDir stempDir
End Sub

Public Sub write_code()
    Text1.Text = "exp amit/ranjan grants=y file=" & Label8.Caption & "\RestBackupFile.DMP"
End Sub

Public Sub final_Run()
    Shell "cmd.exe /c " & Text1.Text
End Sub


Private Sub sub_admin_cbackup_Click()
    Call Browse_Folder
    Call write_code
    Call final_Run
    MsgBox "Backup completed successfully"
End Sub



Private Sub about__Click()
about_company.Show
End Sub

Private Sub admin_cbackup_Click()
 Call Browse_Folder
    Call write_code
    Call final_Run
    MsgBox "Backup completed successfully"
End Sub

Private Sub customer_entry_Click()
customer_entry_form.Show
End Sub

Private Sub customer_order__Click()
CUSTOMER_ORDER.Show
End Sub

Private Sub customer_report_Click()
customer_report_function
End Sub



Private Sub home__Click()
Unload admin
Unload purchased_product
Unload sells
Unload stock_form
Unload report_form
Load m_form
m_form.WindowState = vbMaximized
Picture2.Visible = False
m_form.Show
End Sub

Private Sub Image10_Click()
Unload admin
Unload purchased_product
Unload sells
Unload stock_form
Unload report_form
Load m_form
m_form.WindowState = vbMaximized
Picture2.Visible = False
m_form.Show
End Sub

Private Sub Image11_Click()

Unload customer_entry_form
Unload supplier_entry_form
Unload product_detail
Unload m_form
admin.Show
admin.WindowState = vbMaximized
End Sub

Private Sub Image12_Click()
Unload CUSTOMER_ORDER
Unload supplier_order
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
Unload m_form

order.Show
order.WindowState = vbMaximized
End Sub

Private Sub Image13_Click()
Unload CUSTOMER_ORDER
Unload supplier_order
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
Unload m_form

Load purchased_product
purchased_product.Show
End Sub

Private Sub Image14_Click()
stock_form.WindowState = vbMaximized
Unload m_form
Unload admin
Unload purchased_product
Unload sells
Unload report_form
Load stock_form
stock_form.Show
End Sub

Private Sub Image15_Click()
Unload return_form
Unload m_form
Unload admin
Unload purchased_product
Unload report_form
Unload stock_form
sells.Show
End Sub

Private Sub Image16_Click()
Unload CUSTOMER_ORDER
Unload supplier_order
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
Unload sells
Unload order
return_form.Show
End Sub

Private Sub Image17_Click()
Unload m_form
Unload admin
Unload purchased_product
Unload sells
Unload stock_form
Unload CUSTOMER_ORDER
Unload supplier_order
Unload stock_form
Unload report_form
Unload product_detail
Unload admin
Unload purchased_product
Unload report
Unload sell_return
Unload purchase_return

report_form.Show
End Sub

Private Sub Image18_Click()


If Picture1.Visible = True Then

Picture1.Visible = False
Else
Picture1.Visible = True
End If

If Picture1.Visible = True Then

Timer3.Enabled = True
Else
Timer3.Enabled = False
End If
End Sub

Private Sub Image19_Click()
Unload admin
Unload purchased_product
Unload sells
Unload stock_form
Unload report_form
Load m_form
m_form.WindowState = vbMaximized
Picture2.Visible = False
m_form.Show
End Sub

Private Sub Image9_Click()
ans = MsgBox("Do you Want to Restart Program", vbOKCancel + vbInformation)
If ans = 1 Then
Unload Me
Load MDIForm1
End If
End Sub

Private Sub Label6_Click()
ans = MsgBox("Do you want to log out ", vbYesNo + vbInformation)
If ans = vbYes Then
Unload Me
Load Form2
Form2.Show
End If
End Sub

Private Sub MDIForm_Load()
Load m_form
m_form.Show
Picture1.Visible = False
Image10.MousePointer = vbCustom
Image10.MouseIcon = LoadPicture("C:\Program Files x86\Microsoft Visual Studio\COMMON\Graphics\Cursors\H_POINT.cur")

Image11.MousePointer = vbCustom
Image11.MouseIcon = LoadPicture("C:\Program Files x86\Microsoft Visual Studio\COMMON\Graphics\Cursors\H_POINT.cur")

Image12.MousePointer = vbCustom
Image12.MouseIcon = LoadPicture("C:\Program Files x86\Microsoft Visual Studio\COMMON\Graphics\Cursors\H_POINT.cur")

Image13.MousePointer = vbCustom
Image13.MouseIcon = LoadPicture("C:\Program Files x86\Microsoft Visual Studio\COMMON\Graphics\Cursors\H_POINT.cur")

Image14.MousePointer = vbCustom
Image14.MouseIcon = LoadPicture("C:\Program Files x86\Microsoft Visual Studio\COMMON\Graphics\Cursors\H_POINT.cur")

Image15.MousePointer = vbCustom
Image15.MouseIcon = LoadPicture("C:\Program Files x86\Microsoft Visual Studio\COMMON\Graphics\Cursors\H_POINT.cur")

Image16.MousePointer = vbCustom
Image16.MouseIcon = LoadPicture("C:\Program Files x86\Microsoft Visual Studio\COMMON\Graphics\Cursors\H_POINT.cur")

Image19.MousePointer = vbCustom
Image19.MouseIcon = LoadPicture("C:\Program Files x86\Microsoft Visual Studio\COMMON\Graphics\Cursors\H_POINT.cur")

Image17.MousePointer = vbCustom
Image17.MouseIcon = LoadPicture("C:\Program Files x86\Microsoft Visual Studio\COMMON\Graphics\Cursors\H_POINT.cur")

Label2.Caption = Time
Label3.Caption = Date

Picture2.Visible = False

MDIForm1.Caption = "Dry Fruits Management System"

End Sub



Private Sub order_report_Click()
order_report_function
End Sub

Private Sub product_entry_Click()
product_detail.Show
End Sub

Private Sub product_report_Click()
product_report_function
End Sub

Private Sub purchase__Click()
Unload m_form
Unload admin
Unload sells
Unload report_form
Load purchased_product
purchased_product.Show
End Sub

Private Sub purchase_report_Click()
purchase_report_function
End Sub

Private Sub purchase_return__Click()
Unload sell_return
Unload m_form
MDIForm1.Picture2.Visible = True
purchase_return.Show
purchase_return.WindowState = vbMaximized
End Sub

Private Sub return_report_Click()
return_report_function
End Sub

Private Sub sell__Click()
Unload return_form
Unload m_form
Unload admin
Unload purchased_product
Unload report_form
Unload stock_form
sells.Show
End Sub

Private Sub sell_report_Click()
sale_report_function
End Sub

Private Sub sell_return__Click()
sell_return.Show
End Sub


Private Sub stock__Click()
stock_form.WindowState = vbMaximized
Unload m_form
Unload admin
Unload purchased_product
Unload sells
Unload report_form
Load stock_form
stock_form.Show
End Sub

Private Sub stock_report_Click()
stock_report_function
End Sub

Private Sub supplier_entry_Click()
supplier_entry_form.Show
End Sub


Private Sub supplier_order__Click()
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
supplier_order.WindowState = vbMaximized
End Sub

Private Sub supplier_report_Click()
supplier_report_function
End Sub

Private Sub Timer1_Timer()
Label2.Caption = Time
End Sub

Private Sub Timer3_Timer()
Static hdd As Integer

hdd = hdd + 2
If hdd = 50 Then
Picture1.Visible = False
hdd = 0
End If
End Sub

Private Sub ts_Click()
Help_form.Show
End Sub
