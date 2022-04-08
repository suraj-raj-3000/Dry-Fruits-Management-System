VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Login Form"
   ClientHeight    =   12915
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21915
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   15615
   ScaleWidth      =   28560
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7095
      Left            =   9270
      TabIndex        =   5
      Top             =   2160
      Visible         =   0   'False
      Width           =   5895
      Begin VB.TextBox phone_no 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   10
         Top             =   3600
         Width           =   4335
      End
      Begin VB.CheckBox Check1 
         Height          =   195
         Left            =   840
         TabIndex        =   13
         Top             =   5400
         Width           =   255
      End
      Begin VB.TextBox c_pass 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         TabIndex        =   12
         Top             =   4560
         Width           =   1935
      End
      Begin VB.TextBox pass 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   11
         Top             =   4560
         Width           =   1935
      End
      Begin VB.TextBox email 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   9
         Top             =   2640
         Width           =   4335
      End
      Begin VB.TextBox l_name 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         TabIndex        =   8
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox f_name 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   7
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Phone No :-"
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
         TabIndex        =   25
         Top             =   3240
         Width           =   1230
      End
      Begin VB.Image Image2 
         Height          =   615
         Left            =   2040
         Top             =   5880
         Width           =   1935
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Login Here"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   1
         Left            =   3480
         TabIndex        =   23
         Top             =   6600
         Width           =   1155
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Already have a account"
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
         Index           =   0
         Left            =   960
         TabIndex        =   22
         Top             =   6600
         Width           =   2460
      End
      Begin VB.Label sub1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "SUBMIT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Left            =   2535
         TabIndex        =   14
         Top             =   6030
         Width           =   1020
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000D&
         X1              =   2880
         X2              =   4800
         Y1              =   5640
         Y2              =   5640
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Term and conditions"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   1
         Left            =   2760
         TabIndex        =   21
         Top             =   5400
         Width           =   2130
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "I agree to  the"
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
         Index           =   0
         Left            =   1200
         TabIndex        =   20
         Top             =   5400
         Width           =   1440
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Confirm Password :-"
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
         Left            =   3240
         TabIndex        =   19
         Top             =   4200
         Width           =   2070
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Password :-"
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
         TabIndex        =   18
         Top             =   4200
         Width           =   1230
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Email / Username :-"
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
         TabIndex        =   17
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Last Name :-"
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
         Left            =   3240
         TabIndex        =   16
         Top             =   1320
         Width           =   1320
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "First Name :-"
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
         TabIndex        =   15
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SIGN UP FORM"
         BeginProperty Font 
            Name            =   "Ravie"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   480
         Left            =   1200
         TabIndex        =   6
         Top             =   360
         Width           =   3405
      End
      Begin VB.Shape sub 
         BackColor       =   &H8000000D&
         BorderColor     =   &H8000000D&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C000C0&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   2040
         Top             =   5880
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   7095
      Left            =   15600
      TabIndex        =   24
      Top             =   2160
      Visible         =   0   'False
      Width           =   6015
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   30
         Top             =   4080
         Width           =   4335
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   28
         Top             =   2520
         Width           =   4335
      End
      Begin VB.Line Line3 
         BorderColor     =   &H8000000D&
         X1              =   2160
         X2              =   3870
         Y1              =   6480
         Y2              =   6480
      End
      Begin VB.Label log 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Goto Login page"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Left            =   2160
         TabIndex        =   34
         Top             =   6240
         Width           =   1740
      End
      Begin VB.Image Image3 
         Height          =   975
         Left            =   1800
         Top             =   5400
         Width           =   2175
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "FIND"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   2715
         TabIndex        =   33
         Top             =   5550
         Width           =   660
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Or"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   345
         Left            =   2400
         TabIndex        =   32
         Top             =   3240
         Width           =   390
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No :-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   840
         TabIndex        =   31
         Top             =   3720
         Width           =   1230
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email / Username :-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   840
         TabIndex        =   29
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Line Line2 
         X1              =   2160
         X2              =   3375
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password recovery"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   345
         Left            =   1200
         TabIndex        =   27
         Top             =   1560
         Width           =   2580
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DFMS"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Left            =   2280
         TabIndex        =   26
         Top             =   480
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H8000000D&
         BorderColor     =   &H8000000D&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H8000000D&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   2040
         Top             =   5400
         Width           =   1935
      End
   End
   Begin VB.TextBox user_email 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   4640
      Width           =   4695
   End
   Begin VB.TextBox password 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Top             =   6030
      Width           =   4695
   End
   Begin VB.Image Image1 
      Height          =   1695
      Left            =   4800
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label submit 
      BackStyle       =   0  'Transparent
      Height          =   735
      Left            =   4200
      TabIndex        =   2
      Top             =   7080
      Width           =   3135
   End
   Begin VB.Label signup 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   6480
      TabIndex        =   4
      Top             =   8280
      Width           =   975
   End
   Begin VB.Label forget 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   8280
      Width           =   1935
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public user As String



Private Sub f_name_LostFocus()
f_name.Text = UCase(f_name.Text)
End Sub

Private Sub forget_Click()
Frame1.Visible = False
Frame2.Visible = True
End Sub

Private Sub Form_Load()
Connection
End Sub

Private Sub Frame2_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Image2_Click()
If f_name.Text = "" Or l_name.Text = "" Or Check1.Value <> Checked Or email.Text = "" Or pass.Text = "" Or c_pass.Text = "" Then
MsgBox "fields are blanks"
Else

If pass.Text <> c_pass.Text Then
MsgBox "Password not matched"

Else
user_id
ans = MsgBox("Do you want to Submit ", vbYesNo + vbInformation)
If ans = vbYes Then

Set r = c.Execute("insert into login values('" & user & "','" + f_name.Text + "','" + l_name.Text + "','" + email.Text + "','" + pass.Text + "','" + c_pass.Text + "'," + phone_no.Text + ") ")

MsgBox "User created"
clear
End If

End If
End If
End Sub

Private Sub Image3_Click()
If Text1.Text = "" And Text2.Text = "" Then
 MsgBox "Enter Email or Phone no"
Else
 If Text1.Text <> blank Then
  Set r = c.Execute("select * from login where email='" + Text1.Text + "' ")
  If Text1.Text = r.Fields(3) Then
   a = MsgBox("your Password is " & r.Fields(4), vbCritical, "Password")
  Else
   MsgBox "Email does not Exists"
  End If
  
 ElseIf Text2.Text <> blank Then
   Set r = c.Execute("select * from login where phone_no='" + Text2.Text + "' ")
   If Text2.Text = r.Fields(6) Then
   a = MsgBox("your Password is " & r.Fields(4), vbCritical, "Password")
  Else
   MsgBox "Email does not Exists"
  End If

 End If
 
End If
End Sub

Private Sub l_name_LostFocus()
l_name.Text = UCase(l_name.Text)
End Sub

Private Sub Label10_Click(Index As Integer)
Frame1.Visible = False
Frame2.Visible = False
End Sub



Private Sub Label16_Click()

End Sub

Private Sub log_Click()
Frame1.Visible = False
Frame2.Visible = False
End Sub

Private Sub signup_Click()
Frame1.Visible = True
End Sub

Private Sub submit_Click()
Set r = New ADODB.Recordset
Set r = c.Execute("select * from login where email='" + user_email.Text + "' and password='" + password.Text + "'  ")
If r.EOF = True Then
MsgBox "invalid email or password"
Else
If user_email.Text = r.Fields(3) And password.Text = r.Fields(4) Then

Set r = c.Execute("select f_name,l_name,email from login where email='" + user_email.Text + "' and password='" + password.Text + "'")
MDIForm1.Label4.Caption = r.Fields(0) & " " & r.Fields(1)
'MDIForm1.Label5.Caption = r.Fields(2)
Unload Me
MDIForm1.Show

Else
MsgBox "invalid User Or Password"
End If
End If
End Sub

Public Function user_id()
Dim a As String
Dim i As Integer

Set r = New ADODB.Recordset
sql = "select max(to_number(substr(login_id,7,length(login_id)))) from login"
Set r = c.Execute(sql)

If IsNull(r.Fields(0)) Then
user = "USER" & "00" & 1
Else
user = "USER" & "00" & r.Fields(0) + 1
End If
a = user
If (a = "USER" & "001" & "0") Then
sql = "select max(to_number(substr(LOGIN_ID,4,length(login_id)))) from login"
Set r = c.Execute(sql)
user = "USER" & "0" & r.Fields(0) + 1


End If

End Function

Public Function clear()
f_name.Text = ""
l_name.Text = ""
email.Text = ""
pass.Text = ""
c_pass.Text = ""
Check1.Value = Unchecked
phone_no.Text = ""
End Function
