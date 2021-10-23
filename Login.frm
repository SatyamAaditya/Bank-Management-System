VERSION 5.00
Begin VB.Form Login 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SA Bank"
   ClientHeight    =   10800
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   19110
   DrawMode        =   2  'Blackness
   DrawStyle       =   5  'Transparent
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000A&
   Icon            =   "Login.frx":0000
   LinkTopic       =   "SA Bank"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   2  'Custom
   Picture         =   "Login.frx":10CA
   ScaleHeight     =   720
   ScaleMode       =   0  'User
   ScaleWidth      =   1280
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "David CLM"
         Size            =   11.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   360
      ItemData        =   "Login.frx":9EE3
      Left            =   9360
      List            =   "Login.frx":9EED
      Style           =   2  'Dropdown List
      TabIndex        =   29
      ToolTipText     =   "ChooseType of User"
      Top             =   5400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame5 
      Caption         =   "Confirm Password"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   735
      Left            =   6720
      TabIndex        =   25
      Top             =   7800
      Visible         =   0   'False
      Width           =   3975
      Begin VB.TextBox CNpass 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   26
         Top             =   360
         Visible         =   0   'False
         Width           =   3675
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   735
      Left            =   6720
      TabIndex        =   23
      Top             =   6720
      Visible         =   0   'False
      Width           =   3975
      Begin VB.TextBox NPass 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   240
         PasswordChar    =   "*"
         TabIndex        =   24
         Top             =   360
         Visible         =   0   'False
         Width           =   3705
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Phone"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   6720
      TabIndex        =   17
      Top             =   7800
      Visible         =   0   'False
      Width           =   3975
      Begin VB.TextBox FPhone 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         TabIndex        =   20
         Top             =   200
         Visible         =   0   'False
         Width           =   2775
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "E-Mail"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   6720
      TabIndex        =   16
      Top             =   6720
      Visible         =   0   'False
      Width           =   3975
      Begin VB.TextBox Fmail 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         TabIndex        =   19
         Top             =   210
         Visible         =   0   'False
         Width           =   2775
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   6720
      TabIndex        =   15
      Top             =   5640
      Visible         =   0   'False
      Width           =   3975
      Begin VB.TextBox Fid 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   18
         Top             =   240
         Visible         =   0   'False
         Width           =   2655
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   18120
      Top             =   1320
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   360
      ItemData        =   "Login.frx":9F02
      Left            =   12840
      List            =   "Login.frx":9F04
      MouseIcon       =   "Login.frx":9F06
      MousePointer    =   99  'Custom
      Style           =   2  'Dropdown List
      TabIndex        =   3
      ToolTipText     =   "Select Type of User"
      Top             =   2640
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DownPicture     =   "Login.frx":AFD0
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   10920
      MaskColor       =   &H000000FF&
      MouseIcon       =   "Login.frx":C920
      MousePointer    =   99  'Custom
      Picture         =   "Login.frx":D9EA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7200
      Width           =   3120
   End
   Begin VB.TextBox Password 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      IMEMode         =   3  'DISABLE
      Left            =   9720
      PasswordChar    =   "•"
      TabIndex        =   1
      ToolTipText     =   "Enter Password"
      Top             =   5160
      Width           =   5655
   End
   Begin VB.TextBox Id 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      HideSelection   =   0   'False
      Left            =   9720
      TabIndex        =   0
      ToolTipText     =   "Enter UserName"
      Top             =   3480
      Width           =   5655
   End
   Begin VB.Label info 
      BackStyle       =   0  'Transparent
      Height          =   900
      Left            =   1440
      MouseIcon       =   "Login.frx":F376
      MousePointer    =   99  'Custom
      TabIndex        =   30
      Top             =   7800
      Width           =   896
   End
   Begin VB.Label RPass 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Reset Password"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7510
      TabIndex        =   28
      Top             =   9000
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label FBack 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   5880
      MouseIcon       =   "Login.frx":10440
      MousePointer    =   99  'Custom
      TabIndex        =   27
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Create New Password"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   22
      Top             =   5880
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label FNext 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NEXT"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   7560
      TabIndex        =   21
      Top             =   9030
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Site 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   10920
      MouseIcon       =   "Login.frx":1150A
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   9600
      Width           =   495
   End
   Begin VB.Label Mail 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   9960
      MouseIcon       =   "Login.frx":125D4
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   9600
      Width           =   615
   End
   Begin VB.Label Linkedin 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   9120
      MouseIcon       =   "Login.frx":1369E
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   9600
      Width           =   495
   End
   Begin VB.Label Insta 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   8400
      MouseIcon       =   "Login.frx":14768
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   9600
      Width           =   375
   End
   Begin VB.Label twitter 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   7560
      MouseIcon       =   "Login.frx":15832
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   9600
      Width           =   375
   End
   Begin VB.Label fb 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   6720
      MouseIcon       =   "Login.frx":168FC
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   9720
      Width           =   495
   End
   Begin VB.Label FClose 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   10920
      MouseIcon       =   "Login.frx":179C6
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label FPass 
      BackStyle       =   0  'Transparent
      Caption         =   "Forgotten Password?"
      BeginProperty Font 
         Name            =   "Sanskrit Text"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   13200
      MouseIcon       =   "Login.frx":18A90
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   6360
      Width           =   2295
   End
   Begin VB.Label Msg 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000C000&
      Height          =   615
      Left            =   18480
      MouseIcon       =   "Login.frx":19B5A
      MousePointer    =   99  'Custom
      TabIndex        =   6
      ToolTipText     =   "Secure"
      Top             =   9120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Exit_app 
      BackStyle       =   0  'Transparent
      Height          =   855
      Left            =   17160
      MouseIcon       =   "Login.frx":1AC24
      MousePointer    =   99  'Custom
      TabIndex        =   5
      ToolTipText     =   "Exit"
      Top             =   7800
      Width           =   855
   End
   Begin VB.Label Time 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   15600
      TabIndex        =   4
      Top             =   240
      Width           =   3255
   End
   Begin VB.Image FPassw 
      Height          =   10005
      Left            =   5760
      Picture         =   "Login.frx":1BCEE
      Top             =   240
      Visible         =   0   'False
      Width           =   5700
   End
   Begin VB.Image ErrIco 
      Height          =   615
      Left            =   18360
      MouseIcon       =   "Login.frx":1E19C
      MousePointer    =   99  'Custom
      Picture         =   "Login.frx":1F266
      Stretch         =   -1  'True
      ToolTipText     =   "Error"
      Top             =   9120
      Visible         =   0   'False
      Width           =   597
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public con As New ADODB.Connection
Dim cmd As New ADODB.Command
Dim adm As New ADODB.Recordset
Dim cus As New ADODB.Recordset
Dim edet As String
Dim dbid As String

Dim cn As String
Public Function Make_Connection()
    cn = "Provider=MSDAORA.1;User ID=bank/SABank;Data Source=localhost;Persist Security Info=False"
End Function
Private Sub CNpass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        RPass_Click
    End If
End Sub

Private Sub Command1_Click()
    dbid = UCase(Id.Text)
    Make_Connection
    
    If Combo1.Text = "" Then
        MsgBox "Choose User type", vbExclamation
        Combo1.SetFocus
    Else
        If Id.Text = "" And Password.Text = "" Then
            MsgBox "Enter Username and Password", vbExclamation
            Id.SetFocus
        Else
            
            If Combo1.Text = "Admin" Then
                con.Open cn
                adm.Open " select password from admin_login where id = '" & dbid & "'", cn, adOpenDynamic, adLockOptimistic
                             
                If adm.EOF Then
                    MsgBox "No user found", vbExclamation
                    Id.Text = ""
                    Password.Text = ""
                    adm.Close
                    Id.SetFocus
                Else
                    If adm!Password = Password.Text Then
                        Admin_Home.Show
                        adm.Close
                        Unload Me
                    Else
                        
                        MsgBox "Wrong Password", vbExclamation
                        adm.Close
                        Password.Text = ""
                        Password.SetFocus
                    End If
                End If
                con.Close
            Else
                If Combo1.Text = "Customer" Then
                    
                    con.Open cn
                    cus.Open " select password from customer_data where id = '" & dbid & "'", cn, adOpenDynamic, adLockOptimistic
                    
                    If cus.EOF Then
                        MsgBox "No User Found", vbExclamation
                        cus.Close
                        Id.Text = ""
                        Password.Text = ""
                        Id.SetFocus
                        
                    Else
                        If cus!Password = Password.Text Then
                            Customer_Home.Show
                            cus.Close
                            Unload Me
                            
                        Else
                            MsgBox "Wrong Password", vbExclamation
                            cus.Close
                            Password.Text = ""
                            Password.SetFocus
                        End If
                    End If
                    con.Close
                End If
            End If
        End If
    End If
    
End Sub

Private Sub ErrIco_Click()
MsgBox "Database Error: ,'" & edet & "'", vbCritical
End Sub

Private Sub Exit_app_Click()
Unload Me
End
End Sub

Private Sub fb_Click()
Shell ("explorer https://www.facebook.com/sabankproject")
End Sub

Private Sub FBack_Click()
Label1.Visible = False
Frame4.Visible = False
Frame5.Visible = False
NPass.Visible = False
CNpass.Visible = False
RPass.Visible = False
Frame1.Visible = True
Frame2.Visible = True
Frame3.Visible = True
Fid.Visible = True
Fmail.Visible = True
FPhone.Visible = True
FNext.Visible = True
Combo2.Visible = True
End Sub

Private Sub FClose_Click()
FPassw.Visible = False
Combo1.Visible = True
Command1.Visible = True
Id.Visible = True
Password.Visible = True
FPass.Visible = True
FClose.Visible = False
fb.Visible = True
twitter.Visible = True
Insta.Visible = True
Linkedin.Visible = True
Mail.Visible = True
Site.Visible = True
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Fid.Visible = False
Fmail.Visible = False
FPhone.Visible = False
FNext.Visible = False
Label1.Visible = False
Frame4.Visible = False
Frame5.Visible = False
NPass.Visible = False
CNpass.Visible = False
FBack.Visible = False
RPass.Visible = False
Combo2.Visible = False
Fid.Text = ""
Fmail.Text = ""
FPhone.Text = ""
NPass.Text = ""
CNpass.Text = ""
End Sub

Private Sub FNext_Click()
Dim Fdata As New ADODB.Recordset
Dim Fcdata As New ADODB.Recordset
Dim fdid As String
Dim fdmail As String
Dim fdphone As String
fdid = UCase(Trim(Fid.Text))
If Combo2.Text = "" Then
        MsgBox "Choose User type", vbExclamation
        Combo2.SetFocus
    Else
        If Fid.Text = "" And Fmail.Text = "" And FPhone.Text = "" Then
            MsgBox "Enter ID, Email, and Phone Number", vbExclamation
            Fid.SetFocus
        Else
            
            If Combo2.Text = "Admin" Then
Make_Connection
con.Open cn

Fdata.Open " select * from admin_login where id = '" & fdid & "'", cn, adOpenDynamic, adLockOptimistic

If Fdata.EOF Then
MsgBox "No User Found", vbExclamation
Fid.SetFocus

    Else
    fdmail = UCase(Trim(Fdata!Email))
fdphone = UCase(Trim(Fdata!phone))
    If UCase(Trim(Fmail.Text)) <> fdmail Then
    MsgBox "Wrong E-mail ID", vbExclamation
    Fmail.SetFocus
    Else
    If UCase(Trim(FPhone.Text)) <> fdphone Then
    MsgBox "Wrong Phone Number", vbExclamation
    FPhone.SetFocus
    Else

Label1.Visible = True
Frame4.Visible = True
Frame5.Visible = True
NPass.Visible = True
CNpass.Visible = True
FBack.Visible = True
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Fid.Visible = False
Fmail.Visible = False
FPhone.Visible = False
FNext.Visible = False
Combo2.Visible = False
RPass.Visible = True
End If
End If
End If
con.Close
End If

    If Combo2.Text = "Customer" Then
Make_Connection
con.Open cn

Fcdata.Open " select * from customer_data where id = '" & fdid & "'", cn, adOpenDynamic, adLockOptimistic

If Fcdata.EOF Then
MsgBox "No User Found", vbExclamation
Fid.SetFocus

    Else
    
    If UCase(Trim(Fmail.Text)) <> UCase(Trim(Fcdata!Email)) Then
    MsgBox "Wrong E-mail ID", vbExclamation
    Fmail.SetFocus
    Else
    If UCase(Trim(FPhone.Text)) <> UCase(Trim(Fcdata!Mobile)) Then
    MsgBox "Wrong Phone Number", vbExclamation
    FPhone.SetFocus
    Else

Label1.Visible = True
Frame4.Visible = True
Frame5.Visible = True
NPass.Visible = True
CNpass.Visible = True
FBack.Visible = True
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Fid.Visible = False
Fmail.Visible = False
FPhone.Visible = False
FNext.Visible = False
Combo2.Visible = False
RPass.Visible = True
End If
End If
End If
con.Close
End If
End If
End If

End Sub

Private Sub Form_Activate()
    Combo1.SetFocus
End Sub
Private Sub Form_Load()

Dim test As New ADODB.Recordset
    Combo1.AddItem "Admin"
    Combo1.AddItem "Customer"
    Make_Connection
    
  On Error GoTo Local_Error
test.Open " select ID, PASSWORD, NAME, BRANCH_NAME, ACTYPE, FNAME, MNAME, IFSC, ACNO, DOB,GENDER,MOBILE, EMAIL, ADDRESS, AADHAR, OBALANCE, BALANCE, SIMG, PIMG, LTID, TDEPOSIT, TWITHDRAWL, AODATE, INTERESTRATE, FDPERIOD from customer_data", cn, adOpenDynamic, adLockOptimistic
test.Close
test.Open "select DATES, DEBIT, CREDIT, BALANCE, ID, NAME, ACNO, IFSC, MOBILE, TID, TMODE, TTYPE from customer_statement", cn, adOpenDynamic, adLockOptimistic
test.Close
test.Open " select DATES, TID, TTYPE, TMODE, ID, NAME, AMOUNT, ACNO, IFSC from bank_transaction", cn, adOpenDynamic, adLockOptimistic
test.Close
test.Open "select ID, PASSWORD, NAME, BRANCH_NAME, PHONE, EMAIL, ADDRESS, DOJ, BID, IFSC, TWITTER, LINKEDIN, FACEBOOK, INSTAGRAM, WEBSITE, POST, SALARY, BRANCH_LOCATION, BRANCH_PINCODE, BANK_BAL, TDEPOSIT, TWITHDRAWL from admin_login", cn, adOpenDynamic, adLockOptimistic

Local_Error:
    
    If test.State = adStateOpen Then
    Msg.Visible = True
    test.Close
    Else
    ErrIco.Visible = True
    Msg.Visible = False
    edet = Err.Description
    MsgBox Err.Description, vbCritical, "Oracle Database Error"
    Command1.Enabled = False
    FPass.Enabled = False
    Password.Enabled = False
    End If
    
End Sub

Private Sub FPass_Click()

FPassw.Visible = True
Combo1.Visible = False
Command1.Visible = False
Id.Visible = False
Password.Visible = False
FPass.Visible = False
FClose.Visible = True
fb.Visible = False
twitter.Visible = False
Insta.Visible = False
Linkedin.Visible = False
Mail.Visible = False
Site.Visible = False
Combo2.Visible = True
Frame1.Visible = True
Frame2.Visible = True
Frame3.Visible = True
Fid.Visible = True
Fmail.Visible = True
FPhone.Visible = True
FNext.Visible = True
Label1.Visible = False
Frame4.Visible = False
Frame5.Visible = False
NPass.Visible = False
CNpass.Visible = False
FBack.Visible = False
RPass.Visible = False

End Sub
Private Sub FPhone_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        FNext_Click
    End If
    If KeyAscii <> 48 And KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 52 And KeyAscii <> 53 And KeyAscii <> 54 And KeyAscii <> 55 And KeyAscii <> 56 And KeyAscii <> 57 And KeyAscii <> 8 And KeyAscii <> 13 Then
    KeyAscii = 0
    MsgBox ("Please Enter Numbers Only")
    End If
End Sub

Private Sub FPhone_Validate(Cancel As Boolean)
If Len(FPhone.Text) > 10 Then
      MsgBox "Enter the phone number in 10 digits!", vbExclamation, ""
      Cancel = True
   End If
End Sub

Private Sub info_Click()
MsgBox "Developed By Satyam Aaditya"
Shell ("explorer https://www.satyamaaditya.com/")
End Sub

Private Sub Insta_Click()
Shell ("explorer https://www.instagram.com/sabankproject/")
End Sub

Private Sub Linkedin_Click()
Shell ("explorer https://www.linkedin.com/in/sabankproject")
End Sub

Private Sub Mail_Click()
Shell ("explorer sabankproject@gmail.com")
End Sub

Private Sub Msg_Click()
MsgBox "Securely, Connected With Oracle Database", vbInformation, " Connection"
End Sub

Private Sub Password_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Command1_Click
    End If
End Sub

Private Sub RPass_Click()
Dim gtf As String
Dim cpass As String
Dim gid As String

cpass = CNpass.Text

If Combo2.Text = "Admin" Then
gtf = "admin_login"
Else
If Combo2.Text = "Customer" Then
gtf = "Customer_data"
End If
End If

gid = UCase(Trim(Fid.Text))
If NPass.Text <> CNpass.Text Then
MsgBox ("Password and Cofirmed Password doesn't Match"), vbCritical
Else
If NPass.Text = CNpass.Text Then
Make_Connection
con.Open cn
con.Execute ("update  " & gtf & "  set password='" & cpass & "' where id='" & gid & "' ")
con.Execute ("commit")
MsgBox ("ID " + gid + " Password Changed Successfully")
FClose_Click
Combo1.SetFocus
End If
con.Close
End If
End Sub

Private Sub Site_Click()
Shell ("explorer https://sabankproject.blogspot.com/")
End Sub
Private Sub Timer1_Timer()
    Time.Caption = Now
End Sub
Private Sub twitter_Click()
Shell ("explorer https://www.twitter.com/sabankproject")
End Sub
