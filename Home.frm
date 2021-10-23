VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Admin_Home 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dashboard"
   ClientHeight    =   10800
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   19110
   Icon            =   "Home.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Home.frx":10CA
   ScaleHeight     =   720
   ScaleMode       =   0  'User
   ScaleWidth      =   1280
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView TransactionList 
      Height          =   4200
      Left            =   4080
      TabIndex        =   142
      Top             =   6360
      Visible         =   0   'False
      Width           =   14497
      _ExtentX        =   25559
      _ExtentY        =   7408
      LabelEdit       =   1
      SortOrder       =   -1  'True
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.ListBox Enm7 
      Appearance      =   0  'Flat
      BackColor       =   &H00FAF6F5&
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00378000&
      Height          =   4590
      ItemData        =   "Home.frx":1815D
      Left            =   15621
      List            =   "Home.frx":1815F
      TabIndex        =   102
      Top             =   4950
      Visible         =   0   'False
      Width           =   2732
   End
   Begin VB.ListBox Enm6 
      Appearance      =   0  'Flat
      BackColor       =   &H00FAF6F5&
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   4590
      ItemData        =   "Home.frx":18161
      Left            =   14280
      List            =   "Home.frx":18163
      TabIndex        =   101
      Top             =   4950
      Visible         =   0   'False
      Width           =   1493
   End
   Begin VB.ListBox Enm5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FAF6F5&
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00378000&
      Height          =   4590
      ItemData        =   "Home.frx":18165
      Left            =   11280
      List            =   "Home.frx":18167
      TabIndex        =   100
      Top             =   4950
      Visible         =   0   'False
      Width           =   3180
   End
   Begin VB.ListBox Enm4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FAF6F5&
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   4590
      ItemData        =   "Home.frx":18169
      Left            =   9665
      List            =   "Home.frx":1816B
      TabIndex        =   99
      Top             =   4950
      Visible         =   0   'False
      Width           =   1821
   End
   Begin VB.ListBox Enm3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FAF6F5&
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00378000&
      Height          =   4590
      ItemData        =   "Home.frx":1816D
      Left            =   8256
      List            =   "Home.frx":1816F
      TabIndex        =   98
      Top             =   4950
      Visible         =   0   'False
      Width           =   1612
   End
   Begin VB.ListBox Enm2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FAF6F5&
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   4590
      ItemData        =   "Home.frx":18171
      Left            =   6738
      List            =   "Home.frx":18173
      TabIndex        =   97
      Top             =   4950
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Home.frx":18175
            Key             =   "UpArrow"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Home.frx":1924F
            Key             =   "DownArrow"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView StatementList 
      Height          =   4200
      Left            =   4080
      TabIndex        =   132
      Top             =   6360
      Visible         =   0   'False
      Width           =   14497
      _ExtentX        =   25559
      _ExtentY        =   7408
      View            =   3
      LabelEdit       =   1
      SortOrder       =   -1  'True
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   12582912
      BackColor       =   16447221
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.ComboBox FDPer 
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      ItemData        =   "Home.frx":1A329
      Left            =   15000
      List            =   "Home.frx":1A354
      Style           =   2  'Dropdown List
      TabIndex        =   131
      ToolTipText     =   "Choose FD Period"
      Top             =   1395
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Bindings        =   "Home.frx":1A3BE
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd-mmm-yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   3
      EndProperty
      Height          =   375
      Left            =   11520
      TabIndex        =   120
      Top             =   5640
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16777215
      CustomFormat    =   "dd-MMM-yy"
      Format          =   120193027
      CurrentDate     =   44397
   End
   Begin VB.TextBox DOB 
      BackColor       =   &H00FAF6F5&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd-mmm-yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   3
      EndProperty
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   30
      ToolTipText     =   "Enter Date Of Birth in dd/mm/yyyy Format"
      Top             =   5520
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.TextBox EmpAdd 
      Alignment       =   2  'Center
      BackColor       =   &H00FAF6F5&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15720
      Locked          =   -1  'True
      TabIndex        =   115
      Text            =   " "
      Top             =   4500
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox EmpSal 
      Alignment       =   2  'Center
      BackColor       =   &H00FAF6F5&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14318
      Locked          =   -1  'True
      TabIndex        =   114
      Text            =   " "
      Top             =   4500
      Visible         =   0   'False
      Width           =   1254
   End
   Begin VB.TextBox EmpMail 
      Alignment       =   2  'Center
      BackColor       =   &H00FAF6F5&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11400
      Locked          =   -1  'True
      TabIndex        =   113
      Text            =   " "
      Top             =   4500
      Visible         =   0   'False
      Width           =   2732
   End
   Begin VB.TextBox EmpPhone 
      Alignment       =   2  'Center
      BackColor       =   &H00FAF6F5&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9689
      Locked          =   -1  'True
      TabIndex        =   112
      Text            =   " "
      Top             =   4500
      Visible         =   0   'False
      Width           =   1538
   End
   Begin VB.TextBox EmpEID 
      Alignment       =   2  'Center
      BackColor       =   &H00FAF6F5&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   111
      Text            =   " "
      Top             =   4500
      Visible         =   0   'False
      Width           =   1299
   End
   Begin VB.TextBox EmpDOJ 
      Alignment       =   2  'Center
      BackColor       =   &H00FAF6F5&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   110
      Text            =   " "
      Top             =   4500
      Visible         =   0   'False
      Width           =   1299
   End
   Begin VB.TextBox EmpPosition 
      Alignment       =   2  'Center
      BackColor       =   &H00FAF6F5&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   109
      Text            =   " "
      Top             =   4500
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.TextBox EmpNm 
      Alignment       =   2  'Center
      BackColor       =   &H00FAF6F5&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   108
      Text            =   " "
      Top             =   4500
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox wamt 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10245
      TabIndex        =   90
      Top             =   6780
      Visible         =   0   'False
      Width           =   2687
   End
   Begin VB.TextBox CSigntext 
      Height          =   285
      Left            =   15600
      TabIndex        =   88
      Top             =   4440
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox CPictext 
      Height          =   285
      Left            =   12840
      TabIndex        =   87
      Top             =   4440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CD4 
      Left            =   0
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CD3 
      Left            =   0
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Caadh 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12720
      TabIndex        =   75
      Top             =   9810
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.TextBox CLocation 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12720
      TabIndex        =   74
      Top             =   8790
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.TextBox CMail 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12720
      TabIndex        =   73
      Top             =   7725
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.TextBox CMobile 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12840
      TabIndex        =   72
      Top             =   6690
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.TextBox CGender 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12840
      TabIndex        =   71
      Top             =   5700
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.TextBox CDOB 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   69
      Top             =   9810
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.TextBox CMName 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   67
      Top             =   8760
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.TextBox CFName 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4942
      TabIndex        =   65
      Top             =   7740
      Visible         =   0   'False
      Width           =   5160
   End
   Begin VB.TextBox cname 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4942
      TabIndex        =   63
      Top             =   6720
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.TextBox cac 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4957
      TabIndex        =   62
      Top             =   5700
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.TextBox ccid 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   4957
      TabIndex        =   59
      Top             =   2580
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.TextBox SearchBox 
      BackColor       =   &H00FFFFFF&
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
      Height          =   525
      Left            =   3720
      TabIndex        =   57
      ToolTipText     =   "Enter  ID"
      Top             =   1320
      Visible         =   0   'False
      Width           =   4455
   End
   Begin MSComDlg.CommonDialog CD2 
      Left            =   0
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox ASignText 
      BackColor       =   &H00FAF6F5&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14839
      TabIndex        =   56
      Top             =   7230
      Visible         =   0   'False
      Width           =   3538
   End
   Begin VB.TextBox APimgtext 
      BackColor       =   &H00FAF6F5&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14840
      TabIndex        =   55
      Top             =   5580
      Visible         =   0   'False
      Width           =   3538
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   0
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox OBal 
      BackColor       =   &H00FAF6F5&
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
      Height          =   450
      Left            =   16800
      TabIndex        =   39
      Top             =   2280
      Visible         =   0   'False
      Width           =   1747
   End
   Begin VB.TextBox acno 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15480
      TabIndex        =   38
      Top             =   1320
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Frame Gender 
      BackColor       =   &H00FAF6F5&
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   8880
      TabIndex        =   34
      Top             =   6300
      Visible         =   0   'False
      Width           =   7095
      Begin VB.OptionButton Transgender 
         BackColor       =   &H00FAF6F5&
         Caption         =   "Transgender"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   37
         Top             =   120
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.OptionButton Female 
         BackColor       =   &H00FAF6F5&
         Caption         =   "Female"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   36
         Top             =   120
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.OptionButton Male 
         BackColor       =   &H00FAF6F5&
         Caption         =   "Male"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Top             =   120
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin VB.TextBox Address 
      BackColor       =   &H00FAF6F5&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      ScrollBars      =   1  'Horizontal
      TabIndex        =   33
      ToolTipText     =   "Enter Address"
      Top             =   8850
      Visible         =   0   'False
      Width           =   9615
   End
   Begin VB.TextBox Email 
      BackColor       =   &H00FAF6F5&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   32
      ToolTipText     =   "Enter valid E-mail ID"
      Top             =   8010
      Visible         =   0   'False
      Width           =   9615
   End
   Begin VB.TextBox Mobile 
      BackColor       =   &H00FAF6F5&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   31
      ToolTipText     =   "Enter 10 Digit Mobile Number"
      Top             =   7200
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.TextBox MName 
      BackColor       =   &H00FAF6F5&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   29
      ToolTipText     =   "Enter Mother's Name"
      Top             =   4680
      Visible         =   0   'False
      Width           =   9615
   End
   Begin VB.TextBox OName 
      BackColor       =   &H00FAF6F5&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   27
      ToolTipText     =   "Enter Name"
      Top             =   3000
      Visible         =   0   'False
      Width           =   9615
   End
   Begin VB.OptionButton FD 
      BackColor       =   &H00FAF6F5&
      Caption         =   "FD"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13440
      TabIndex        =   26
      Top             =   1350
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.OptionButton Current 
      BackColor       =   &H00FAF6F5&
      Caption         =   "Current"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11280
      TabIndex        =   25
      Top             =   1350
      Visible         =   0   'False
      Width           =   1911
   End
   Begin VB.OptionButton Saving 
      BackColor       =   &H00FAF6F5&
      Caption         =   "Saving"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      TabIndex        =   24
      Top             =   1350
      Visible         =   0   'False
      Width           =   1911
   End
   Begin VB.Timer Date 
      Interval        =   1
      Left            =   0
      Top             =   120
   End
   Begin VB.TextBox CId 
      BackColor       =   &H00FAF6F5&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   23
      Top             =   2220
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.TextBox ctype 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4957
      TabIndex        =   60
      Top             =   3600
      Visible         =   0   'False
      Width           =   5405
   End
   Begin VB.TextBox FName 
      BackColor       =   &H00FAF6F5&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   28
      ToolTipText     =   "Enter Father's Name"
      Top             =   3840
      Visible         =   0   'False
      Width           =   9615
   End
   Begin VB.ListBox Enm1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FAF6F5&
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00378000&
      Height          =   4590
      ItemData        =   "Home.frx":1A3C9
      Left            =   5697
      List            =   "Home.frx":1A3CB
      TabIndex        =   96
      Top             =   4950
      Visible         =   0   'False
      Width           =   1224
   End
   Begin VB.TextBox cifsc 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4957
      TabIndex        =   61
      Top             =   4650
      Visible         =   0   'False
      Width           =   5655
   End
   Begin VB.ListBox Enm 
      Appearance      =   0  'Flat
      BackColor       =   &H00FAF6F5&
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   4590
      ItemData        =   "Home.frx":1A3CD
      Left            =   3930
      List            =   "Home.frx":1A3D4
      TabIndex        =   95
      Top             =   4950
      Visible         =   0   'False
      Width           =   1971
   End
   Begin VB.Label SearchTran 
      BackStyle       =   0  'Transparent
      Height          =   525
      Left            =   8400
      MouseIcon       =   "Home.frx":1A3DD
      MousePointer    =   99  'Custom
      TabIndex        =   143
      Top             =   1275
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Sloc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15915
      TabIndex        =   141
      Top             =   4980
      Visible         =   0   'False
      Width           =   2389
   End
   Begin VB.Label Sintr 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13227
      TabIndex        =   140
      Top             =   4980
      Visible         =   0   'False
      Width           =   2240
   End
   Begin VB.Label Sifs 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10920
      TabIndex        =   139
      Top             =   4980
      Visible         =   0   'False
      Width           =   2299
   End
   Begin VB.Label Scod 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8137
      TabIndex        =   138
      Top             =   4980
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Sbrnm 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   137
      Top             =   4980
      Visible         =   0   'False
      Width           =   2299
   End
   Begin VB.Label SAcNo 
      Alignment       =   2  'Center
      BackColor       =   &H00FAF6F5&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4494
      TabIndex        =   126
      Top             =   4980
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Sttra 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   16080
      TabIndex        =   136
      Top             =   3255
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Stwid 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   12240
      TabIndex        =   135
      ToolTipText     =   "Sum of all Withdrawl Transaction from Bank Opening to Today of All Branch."
      Top             =   3285
      Visible         =   0   'False
      Width           =   2404
   End
   Begin VB.Label Stdep 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   8400
      TabIndex        =   134
      ToolTipText     =   "Sum of all Deposits Transaction from Bank Opening to Today of All Branch."
      Top             =   3285
      Visible         =   0   'False
      Width           =   2419
   End
   Begin VB.Label BankBal 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   4560
      TabIndex        =   133
      ToolTipText     =   "This Amount is all Branch Balance"
      Top             =   3255
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label SNom 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15960
      TabIndex        =   130
      Top             =   4980
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label SIra 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13440
      TabIndex        =   129
      Top             =   4980
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label SAIfsc 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10920
      TabIndex        =   128
      Top             =   4980
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label SAtyp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   127
      Top             =   4980
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label SCbal 
      BackColor       =   &H00FAF6F5&
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   16080
      TabIndex        =   125
      Top             =   3285
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label STWd 
      BackColor       =   &H00FAF6F5&
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12240
      TabIndex        =   124
      Top             =   3285
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label STDp 
      BackColor       =   &H00FAF6F5&
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   123
      Top             =   3285
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label SObal 
      BackColor       =   &H00FAF6F5&
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   122
      Top             =   3285
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label SearchS 
      BackStyle       =   0  'Transparent
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "dd-mm-yyyy HH24:MI:SS"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   0
      EndProperty
      Height          =   735
      Left            =   8160
      MouseIcon       =   "Home.frx":1B4A7
      MousePointer    =   99  'Custom
      TabIndex        =   121
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label EmpSave 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   17760
      MouseIcon       =   "Home.frx":1C571
      MousePointer    =   99  'Custom
      TabIndex        =   119
      Top             =   1695
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.Label EmpDel 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   17040
      MouseIcon       =   "Home.frx":1D63B
      MousePointer    =   99  'Custom
      TabIndex        =   118
      Top             =   1695
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.Label EmpNew 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   15600
      MouseIcon       =   "Home.frx":1E705
      MousePointer    =   99  'Custom
      TabIndex        =   117
      Top             =   1695
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label EmpEdit 
      BackColor       =   &H000000C0&
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   16333
      MouseIcon       =   "Home.frx":1F7CF
      MousePointer    =   99  'Custom
      TabIndex        =   116
      Top             =   1695
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label SEmpBtn 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   8280
      MouseIcon       =   "Home.frx":20899
      MousePointer    =   99  'Custom
      TabIndex        =   107
      Top             =   1080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label TBD3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13885
      TabIndex        =   106
      Top             =   2940
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label TBD2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10943
      TabIndex        =   105
      Top             =   2940
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label TBD1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8062
      TabIndex        =   104
      Top             =   2940
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label TBD 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4583
      TabIndex        =   103
      Top             =   2940
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label DepositBTN 
      BackStyle       =   0  'Transparent
      Height          =   735
      Left            =   10320
      MouseIcon       =   "Home.frx":21963
      MousePointer    =   99  'Custom
      TabIndex        =   94
      Top             =   7920
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label WithdrawBTN 
      BackStyle       =   0  'Transparent
      Height          =   735
      Left            =   10320
      MouseIcon       =   "Home.frx":22A2D
      MousePointer    =   99  'Custom
      TabIndex        =   93
      Top             =   7920
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Wbtn 
      BackStyle       =   0  'Transparent
      Height          =   660
      Left            =   10272
      TabIndex        =   92
      Top             =   7950
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Bal 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14556
      TabIndex        =   91
      Top             =   5625
      Visible         =   0   'False
      Width           =   2986
   End
   Begin VB.Label SCustom 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   8361
      MouseIcon       =   "Home.frx":23AF7
      MousePointer    =   99  'Custom
      TabIndex        =   89
      Top             =   1275
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label CUpdate 
      BackStyle       =   0  'Transparent
      Height          =   855
      Left            =   14400
      TabIndex        =   85
      Top             =   1080
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label CNext 
      BackStyle       =   0  'Transparent
      Height          =   855
      Left            =   12000
      TabIndex        =   84
      Top             =   1080
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label CPrev 
      BackStyle       =   0  'Transparent
      Height          =   855
      Left            =   9480
      TabIndex        =   83
      Top             =   1080
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label CSignedit 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   18360
      MouseIcon       =   "Home.frx":24BC1
      MousePointer    =   99  'Custom
      TabIndex        =   82
      Top             =   4320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label CPicedit 
      BackStyle       =   0  'Transparent
      Height          =   420
      Left            =   14520
      MouseIcon       =   "Home.frx":25C8B
      MousePointer    =   99  'Custom
      TabIndex        =   81
      Top             =   4440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image CSign 
      Enabled         =   0   'False
      Height          =   1815
      Left            =   15480
      Stretch         =   -1  'True
      Top             =   2520
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Image CPic 
      Enabled         =   0   'False
      Height          =   1860
      Left            =   12600
      Stretch         =   -1  'True
      Top             =   2520
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Caadedit 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   18360
      MouseIcon       =   "Home.frx":26D55
      MousePointer    =   99  'Custom
      TabIndex        =   80
      Top             =   9720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label CLoedit 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   18360
      MouseIcon       =   "Home.frx":27E1F
      MousePointer    =   99  'Custom
      TabIndex        =   79
      Top             =   8760
      Visible         =   0   'False
      Width           =   463
   End
   Begin VB.Label CMaedit 
      BackStyle       =   0  'Transparent
      Height          =   405
      Left            =   18360
      MouseIcon       =   "Home.frx":28EE9
      MousePointer    =   99  'Custom
      TabIndex        =   78
      Top             =   7800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label CMoedit 
      BackStyle       =   0  'Transparent
      Height          =   525
      Left            =   18360
      MouseIcon       =   "Home.frx":29FB3
      MousePointer    =   99  'Custom
      TabIndex        =   77
      Top             =   6600
      Visible         =   0   'False
      Width           =   478
   End
   Begin VB.Label CGedit 
      BackStyle       =   0  'Transparent
      Height          =   420
      Left            =   18364
      MouseIcon       =   "Home.frx":2B07D
      MousePointer    =   99  'Custom
      TabIndex        =   76
      Top             =   5640
      Visible         =   0   'False
      Width           =   388
   End
   Begin VB.Label CDedit 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   10320
      MouseIcon       =   "Home.frx":2C147
      MousePointer    =   99  'Custom
      TabIndex        =   70
      Top             =   9720
      Width           =   495
   End
   Begin VB.Label CMedit 
      BackStyle       =   0  'Transparent
      Height          =   420
      Left            =   10316
      MouseIcon       =   "Home.frx":2D211
      MousePointer    =   99  'Custom
      TabIndex        =   68
      Top             =   8700
      Visible         =   0   'False
      Width           =   388
   End
   Begin VB.Label CFedit 
      BackStyle       =   0  'Transparent
      Height          =   450
      Left            =   10320
      MouseIcon       =   "Home.frx":2E2DB
      MousePointer    =   99  'Custom
      TabIndex        =   66
      Top             =   7650
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label CNedit 
      BackStyle       =   0  'Transparent
      Height          =   405
      Left            =   10272
      MouseIcon       =   "Home.frx":2F3A5
      MousePointer    =   99  'Custom
      TabIndex        =   64
      Top             =   6630
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label SearchCustomer 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   8280
      MouseIcon       =   "Home.frx":3046F
      MousePointer    =   99  'Custom
      TabIndex        =   58
      Top             =   1200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label UploadPImg 
      BackStyle       =   0  'Transparent
      Height          =   2535
      Left            =   3480
      TabIndex        =   53
      ToolTipText     =   "Click Here To Upload"
      Top             =   1680
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label ASignUpload 
      BackStyle       =   0  'Transparent
      Height          =   1695
      Left            =   3600
      TabIndex        =   54
      ToolTipText     =   "Click Here To Upload"
      Top             =   5280
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Image ASignImg 
      Height          =   975
      Left            =   3434
      Stretch         =   -1  'True
      Top             =   5250
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Image APImg 
      Height          =   1890
      Left            =   3538
      Stretch         =   -1  'True
      Top             =   1725
      Visible         =   0   'False
      Width           =   2090
   End
   Begin VB.Label ACancel 
      BackStyle       =   0  'Transparent
      Height          =   705
      Left            =   11520
      TabIndex        =   52
      Top             =   9600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Website 
      BackColor       =   &H00F4F2F1&
      Caption         =   "project.satyamaaditya.com"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   16200
      TabIndex        =   50
      Top             =   9975
      Visible         =   0   'False
      Width           =   2448
   End
   Begin VB.Label Instagram 
      BackColor       =   &H00F4F2F1&
      Caption         =   "/SABankProject"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   16200
      TabIndex        =   49
      Top             =   9360
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label facebook 
      BackColor       =   &H00F4F2F1&
      Caption         =   "/SABankProject"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   16200
      TabIndex        =   48
      Top             =   8760
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Linkedin 
      BackColor       =   &H00F4F2F1&
      Caption         =   "/SABankProject"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   16200
      TabIndex        =   47
      Top             =   8205
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Download2 
      Caption         =   "• FormNo. 60"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   3720
      MouseIcon       =   "Home.frx":31539
      MousePointer    =   99  'Custom
      TabIndex        =   42
      ToolTipText     =   "Click Here To Download"
      Top             =   7440
      Visible         =   0   'False
      Width           =   4568
   End
   Begin VB.Label Download1 
      BackColor       =   &H00F4F2F1&
      Caption         =   "• Account Opening Form"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   3720
      MouseIcon       =   "Home.frx":32603
      MousePointer    =   99  'Custom
      TabIndex        =   41
      ToolTipText     =   "Click Here To Download"
      Top             =   6840
      Visible         =   0   'False
      Width           =   4568
   End
   Begin VB.Label Hifsc 
      BackColor       =   &H00F4F2F1&
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4464
      TabIndex        =   40
      Top             =   5115
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Label LOpen 
      Alignment       =   2  'Center
      BackColor       =   &H007C3503&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14900
      TabIndex        =   22
      Top             =   9600
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Label LogNo 
      BackStyle       =   0  'Transparent
      Height          =   1260
      Left            =   12480
      MouseIcon       =   "Home.frx":336CD
      MousePointer    =   99  'Custom
      TabIndex        =   21
      Top             =   6900
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label LogYes 
      BackStyle       =   0  'Transparent
      Height          =   1260
      Left            =   6960
      MouseIcon       =   "Home.frx":34797
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   6900
      Visible         =   0   'False
      Width           =   2762
   End
   Begin VB.Label LName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   15615
      TabIndex        =   14
      Top             =   1260
      Visible         =   0   'False
      Width           =   2520
      WordWrap        =   -1  'True
   End
   Begin VB.Label Paddress 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   16139
      TabIndex        =   19
      Top             =   6960
      Visible         =   0   'False
      Width           =   2055
      WordWrap        =   -1  'True
   End
   Begin VB.Label Pmail 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   16050
      TabIndex        =   18
      Top             =   5805
      Visible         =   0   'False
      Width           =   2055
      WordWrap        =   -1  'True
   End
   Begin VB.Label Pphone 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   16050
      TabIndex        =   17
      Top             =   4635
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label Pdoj 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   16140
      TabIndex        =   16
      Top             =   3510
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Label Pid 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   16080
      TabIndex        =   15
      Top             =   2310
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Image PImg 
      Height          =   1200
      Left            =   16245
      Picture         =   "Home.frx":35861
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label Cross 
      Alignment       =   2  'Center
      BackColor       =   &H00AA4A04&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   660
      Left            =   18330
      TabIndex        =   12
      ToolTipText     =   "Close"
      Top             =   195
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label LogoutPage 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   480
      MouseIcon       =   "Home.frx":369A0
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   9840
      Width           =   2415
   End
   Begin VB.Label EmployeePage 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   480
      MouseIcon       =   "Home.frx":37A6A
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   8880
      Width           =   2415
   End
   Begin VB.Label TransactionPage 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   480
      MouseIcon       =   "Home.frx":38B34
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   7920
      Width           =   2415
   End
   Begin VB.Label StatementPage 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   480
      MouseIcon       =   "Home.frx":39BFE
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   6960
      Width           =   2415
   End
   Begin VB.Label DepositPage 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   480
      MouseIcon       =   "Home.frx":3ACC8
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   5760
      Width           =   2415
   End
   Begin VB.Label WithdrawPage 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   480
      MouseIcon       =   "Home.frx":3BD92
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   4680
      Width           =   2415
   End
   Begin VB.Label CustomerPage 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   480
      MouseIcon       =   "Home.frx":3CE5C
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Label OpenAccountPage 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   480
      MouseIcon       =   "Home.frx":3DF26
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Date_Time 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   300
      Width           =   2535
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   8970
      TabIndex        =   2
      Top             =   240
      Width           =   3735
   End
   Begin VB.Label HomePage 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   480
      MouseIcon       =   "Home.frx":3EFF0
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Image Profile 
      Height          =   7500
      Left            =   14700
      Picture         =   "Home.frx":400BA
      Top             =   120
      Visible         =   0   'False
      Width           =   4275
   End
   Begin VB.Label OpenProfile 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   16305
      MouseIcon       =   "Home.frx":42993
      MousePointer    =   99  'Custom
      TabIndex        =   13
      ToolTipText     =   "Open Profile"
      Top             =   210
      Width           =   2640
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   " "
      DragMode        =   1  'Automatic
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   15240
      TabIndex        =   0
      Top             =   300
      UseMnemonic     =   0   'False
      Width           =   3255
   End
   Begin VB.Label Hmail 
      BackColor       =   &H00F4F2F1&
      Caption         =   "aadityasatyam@gmail.com"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   16200
      TabIndex        =   44
      Top             =   6480
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Hmbl 
      Caption         =   "7909064239"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   16200
      TabIndex        =   43
      Top             =   5880
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Haddress 
      BackColor       =   &H00F4F2F1&
      Caption         =   "Patna, India"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   16200
      TabIndex        =   45
      Top             =   7080
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Twitter 
      BackColor       =   &H00F4F2F1&
      Caption         =   "/SABankProject"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   16200
      TabIndex        =   46
      Top             =   7650
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label BCode 
      BackColor       =   &H00F4F2F1&
      Caption         =   "XXXXX"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14640
      TabIndex        =   51
      Top             =   5145
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image HomeCover 
      Height          =   3900
      Left            =   3150
      Picture         =   "Home.frx":43A5D
      Top             =   780
      Visible         =   0   'False
      Width           =   16035
   End
   Begin VB.Label CDel 
      BackStyle       =   0  'Transparent
      Height          =   855
      Left            =   16920
      TabIndex        =   86
      Top             =   1080
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Image TransactionImg 
      Height          =   10800
      Left            =   0
      Picture         =   "Home.frx":64393
      Top             =   0
      Visible         =   0   'False
      Width           =   19200
   End
   Begin VB.Image StatementImg 
      Height          =   10800
      Left            =   0
      Picture         =   "Home.frx":70D46
      Top             =   0
      Visible         =   0   'False
      Width           =   19200
   End
   Begin VB.Image DepositImg 
      Height          =   10800
      Left            =   0
      Picture         =   "Home.frx":7D0EE
      Top             =   0
      Visible         =   0   'False
      Width           =   19200
   End
   Begin VB.Image HomeImg 
      Height          =   10800
      Left            =   0
      Picture         =   "Home.frx":88AE1
      Top             =   0
      Visible         =   0   'False
      Width           =   19200
   End
   Begin VB.Image PLogout 
      Height          =   9750
      Left            =   0
      Picture         =   "Home.frx":9FB74
      Top             =   1050
      Visible         =   0   'False
      Width           =   19200
   End
   Begin VB.Image AccountImg 
      Height          =   10800
      Left            =   0
      Picture         =   "Home.frx":A9EAA
      Top             =   0
      Visible         =   0   'False
      Width           =   19200
   End
   Begin VB.Image CustomerImg 
      Height          =   10800
      Left            =   0
      Picture         =   "Home.frx":E943A
      Top             =   0
      Visible         =   0   'False
      Width           =   19200
   End
   Begin VB.Image WithdrawImg 
      Height          =   10800
      Left            =   0
      Picture         =   "Home.frx":122EF7
      Top             =   0
      Visible         =   0   'False
      Width           =   19200
   End
   Begin VB.Image EmployeeImg 
      Height          =   10800
      Left            =   0
      Picture         =   "Home.frx":12EC77
      Top             =   0
      Visible         =   0   'False
      Width           =   19200
   End
End
Attribute VB_Name = "Admin_Home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim cmd As New ADODB.Command
Dim rs As New ADODB.Recordset
Dim zo As New ADODB.Recordset
Dim banbal As New ADODB.Recordset
Dim bbala As String
Private Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Dim cn As String
Public Function Make_Connection()
    cn = "Provider=MSDAORA.1;User ID=bank/SABank;Data Source=localhost;Persist Security Info=False"
End Function
Public Sub PlaySound(strFileName As String)
sndPlaySound strFileName, 1
End Sub
Private Sub SortListView(ByRef List As ListView, ColHeadIndex As Integer)
    
    Dim lcv As Long     'Loop Control Variable
  
    With List
        ' Make sure the Sorted property is set to true
        .Sorted = True
        
        ' Sort according to the colum that was clicked (off by one)
        .SortKey = ColHeadIndex - 1
       
        ' Does the column already have an icon?
        If .ColumnHeaders(ColHeadIndex).Icon = 0 Then
            'No, So we will assume this column is not sorted
            
            ' Set to Ascending order
            .SortOrder = lvwAscending
            
            ' Set the ColumnHeader to be the Up Arrow
            .ColumnHeaders(ColHeadIndex).Icon = "UpArrow"
            
        ' Does the column have an UpArrow icon?
        ElseIf .ColumnHeaders(ColHeadIndex).Icon = "UpArrow" Then
            ' Yes, So the column is in Ascending order, switch to descending
            
            ' Set the Column Icon to the Down Arrow
            .ColumnHeaders(ColHeadIndex).Icon = "DownArrow"
            
            ' Set the sort order to descending
            .SortOrder = lvwDescending
        
        Else
            ' Otherwise sort into ascending order
        
            ' Set to Ascending order
            .SortOrder = lvwAscending
            
            ' Set the ColumnHeader to be the Up Arrow
            .ColumnHeaders(ColHeadIndex).Icon = "UpArrow"
        End If
       
        ' Remove any icon (presumably an arrow icon) from all other columns
        ' For every Column in the ListView Control...
        For lcv = 1 To List.ColumnHeaders.Count
            ' Is the current column the clicked column?
            If Not (lcv = ColHeadIndex) Then
                ' No, remove any icon it may have
                .ColumnHeaders(lcv).Icon = 0
            End If
        Next lcv
    
        ' Refresh the display of the ListView Control
        .Refresh
    End With
End Sub
Private Sub ACancel_Click()
PLogout.Visible = False
HomeCover.Visible = True
Hifsc.Visible = True
BCode.Visible = True
Hmbl.Visible = True
Hmail.Visible = True
Haddress.Visible = True
Twitter.Visible = True
Linkedin.Visible = True
facebook.Visible = True
Instagram.Visible = True
Website.Visible = True
Download1.Visible = True
Download2.Visible = True

FDPer.Visible = False
Profile.Visible = False
Cross.Visible = False
LName.Visible = False
Pid.Visible = False
Pdoj.Visible = False
Pphone.Visible = False
Pmail.Visible = False
Paddress.Visible = False
PImg.Visible = False

SCustom.Visible = False
Bal.Visible = False
wamt.Visible = False
Wbtn.Visible = False

Saving.Visible = False
Current.Visible = False
FD.Visible = False
CID.Visible = False
OBal.Visible = False
OName.Visible = False
FName.Visible = False
MName.Visible = False
DOB.Visible = False
DTPicker1.Visible = False
Gender.Visible = False
Male.Visible = False
Female.Visible = False
Transgender.Visible = False
Mobile.Visible = False
Email.Visible = False
Address.Visible = False
PLogout.Visible = False
LogYes.Visible = False
LogNo.Visible = False
APImg.Visible = False
UploadPImg.Visible = False
ASignImg.Visible = False
ASignUpload.Visible = False
APimgtext.Visible = False
ASignText.Visible = False

Enm.Visible = False
Enm1.Visible = False
Enm2.Visible = False
Enm3.Visible = False
Enm4.Visible = False
Enm5.Visible = False
Enm6.Visible = False
Enm7.Visible = False
TBD.Visible = False
TBD1.Visible = False
TBD2.Visible = False
TBD3.Visible = False

SEmpBtn.Visible = False
EmpNm.Visible = False
EmpPosition.Visible = False
EmpDOJ.Visible = False
EmpEID.Visible = False
EmpPhone.Visible = False
EmpMail.Visible = False
EmpSal.Visible = False
EmpAdd.Visible = False

ACNo.Text = ""
CID.Text = ""
OBal.Text = ""
OName.Text = ""
FName.Text = ""
MName.Text = ""
DOB.Text = ""
Mobile.Text = ""
Email.Text = ""
Address.Text = ""
APimgtext.Text = ""
ASignText.Text = ""
APImg.Picture = Nothing
ASignImg.Picture = Nothing

Set Me.Picture = Me.HomeImg
End Sub

Private Sub ASignUpload_Click()
CD2.Filter = "JPG File | *.jpg|GIF File|*.gif|All files|*.*"
CD2.ShowOpen
If CD2.FileName <> "" Then
ASignImg.Picture = LoadPicture(CD2.FileName)
ASignText.Text = CD2.FileName
End If

End Sub
Private Sub Caadedit_Click()
Caadh.Enabled = True
cname.Enabled = False
CFName.Enabled = False
CMName.Enabled = False
CDOB.Enabled = False
CGender.Enabled = False
CMobile.Enabled = False
CMail.Enabled = False
CLocation.Enabled = False
End Sub

Private Sub CDedit_Click()
CDOB.Enabled = False
cname.Enabled = False
CFName.Enabled = False
CMName.Enabled = False
CGender.Enabled = False
CMobile.Enabled = False
CMail.Enabled = False
CLocation.Enabled = False
Caadh.Enabled = False
End Sub

Private Sub CDel_Click()
Dim cdids As String
cdids = ccid.Text
If ccid.Text = "" Then
MsgBox ("Search Customer First")
Else
answer = MsgBox("Are you sure you want to close account of " & vbCrLf & "'" & cname.Text & "'", vbCritical + vbYesNo, "Warning")
If answer = vbYes Then
SCustom_Click
wamt.Text = Bal.Caption
WithdrawBTN_Click
Make_Connection
con.Open cn
con.Execute ("delete from customer_data where id='" & cdids & "' ")
con.Execute ("delete from customer_statement where id='" & cdids & "' ")
con.Execute ("commit")
MsgBox ("A/C No. " + cac.Text + " closed Successfully")

ccid.Text = ""
ctype.Text = ""
cifsc.Text = ""
cac.Text = ""
cname.Text = ""
CFName.Text = ""
CMName.Text = ""
CDOB.Text = ""
CGender.Text = ""
CMobile.Text = ""
CMail.Text = ""
CLocation.Text = ""
Caadh.Text = ""
CPic.Picture = Nothing
CSign.Picture = Nothing
con.Close
End If
End If
End Sub

Private Sub CFedit_Click()
CFName.Enabled = True
cname.Enabled = False
CMName.Enabled = False
CDOB.Enabled = False
CGender.Enabled = False
CMobile.Enabled = False
CMail.Enabled = False
CLocation.Enabled = False
Caadh.Enabled = False
End Sub

Private Sub CGedit_Click()
CGender.Enabled = True
cname.Enabled = False
CFName.Enabled = False
CMName.Enabled = False
CDOB.Enabled = False
CMobile.Enabled = False
CMail.Enabled = False
CLocation.Enabled = False
Caadh.Enabled = False
End Sub

Private Sub CLoedit_Click()
CLocation.Enabled = True
cname.Enabled = False
CFName.Enabled = False
CMName.Enabled = False
CDOB.Enabled = False
CGender.Enabled = False
CMobile.Enabled = False
CMail.Enabled = False
Caadh.Enabled = False
End Sub

Private Sub CMaedit_Click()
CMail.Enabled = True
cname.Enabled = False
CFName.Enabled = False
CMName.Enabled = False
CDOB.Enabled = False
CGender.Enabled = False
CMobile.Enabled = False
CLocation.Enabled = False
Caadh.Enabled = False
End Sub

Private Sub CMedit_Click()
CMName.Enabled = True
cname.Enabled = False
CFName.Enabled = False
CDOB.Enabled = False
CGender.Enabled = False
CMobile.Enabled = False
CMail.Enabled = False
CLocation.Enabled = False
Caadh.Enabled = False
End Sub

Private Sub CMobile_KeyPress(KeyAscii As Integer)
If KeyAscii <> 48 And KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 52 And KeyAscii <> 53 And KeyAscii <> 54 And KeyAscii <> 55 And KeyAscii <> 56 And KeyAscii <> 57 And KeyAscii <> 8 And KeyAscii <> 13 Then
    KeyAscii = 0
    MsgBox ("Please Enter Numbers Only")
    End If
End Sub

Private Sub CMoedit_Click()
CMobile.Enabled = True
cname.Enabled = False
CFName.Enabled = False
CMName.Enabled = False
CDOB.Enabled = False
CGender.Enabled = False
CMail.Enabled = False
CLocation.Enabled = False
Caadh.Enabled = False
End Sub

Private Sub CNedit_Click()
cname.Enabled = True
CFName.Enabled = False
CMName.Enabled = False
CDOB.Enabled = False
CGender.Enabled = False
CMobile.Enabled = False
CMail.Enabled = False
CLocation.Enabled = False
Caadh.Enabled = False

End Sub

Private Sub CNext_Click()
Dim srn As New ADODB.Recordset
Dim snre As String
Dim Snre1 As String
Dim Snre2 As String
cname.Enabled = False
CFName.Enabled = False
CMName.Enabled = False
CDOB.Enabled = False
CGender.Enabled = False
CMobile.Enabled = False
CMail.Enabled = False
CLocation.Enabled = False
Caadh.Enabled = False

rs.Open
Make_Connection
con.Open cn

srn.Open " select * from customer_data order by id", cn, adOpenDynamic, adLockOptimistic
If ccid.Text = "" Then
MsgBox ("Please Search any Record")
Else
snre = Trim(Right(ccid.Text, 11)) + 1
    Snre1 = rs!bid + snre
f = 0
    Do While srn.EOF <> True
    
    If Snre1 = srn!Id Then
    f = 1
ccid.Text = srn!Id
ctype.Text = srn!actype
cifsc.Text = srn!IFSC
cac.Text = srn!ACNo
cname.Text = srn!Name
CFName.Text = srn!FName
CMName.Text = srn!MName
CDOB.Text = srn!DOB
CGender.Text = srn!Gender
CMobile.Text = srn!Mobile
CMail.Text = srn!Email
CLocation.Text = srn!Address
CPic.Picture = LoadPicture(srn!PImg)
CPictext.Text = srn!PImg
CSign.Picture = LoadPicture(srn!sImg)
CSigntext.Text = srn!sImg
SearchBox.Text = ccid.Text
If IsNull(srn!aadhar) Then
Caadh.Text = ""
MsgBox "KYC Pending, Update Aadhar Data", vbInformation
Caadh.Enabled = True
Caadh.SetFocus
Else
Caadh.Text = srn!aadhar
End If
Exit Do
End If
srn.MoveNext
Loop
If f = 0 Then
Snre1 = (rs!bid & (snre - 1))
If ccid.Text = Snre1 Then
MsgBox "Serially Last Record, Search other by ID", vbInformation
Else
If ccid.Text <> Snre1 Then
MsgBox "You Can't See Next Record of other branch, but you can search individually", vbInformation
End If
End If
End If
End If
con.Close
rs.Close
End Sub

Private Sub CPicedit_Click()
cname.Enabled = False
CFName.Enabled = False
CMName.Enabled = False
CDOB.Enabled = False
CGender.Enabled = False
CMobile.Enabled = False
CMail.Enabled = False
CLocation.Enabled = False
Caadh.Enabled = False
CD3.Filter = "JPG File | *.jpg|GIF File|*.gif"
CD3.ShowOpen
If CD3.FileName <> "" Then
CPic.Picture = LoadPicture(CD3.FileName)
CPictext.Text = CD3.FileName
End If
End Sub

Private Sub CPrev_Click()
Dim srp As New ADODB.Recordset
Dim spre As String
Dim spre1 As String
cname.Enabled = False
CFName.Enabled = False
CMName.Enabled = False
CDOB.Enabled = False
CGender.Enabled = False
CMobile.Enabled = False
CMail.Enabled = False
CLocation.Enabled = False
Caadh.Enabled = False

rs.Open
Make_Connection
con.Open cn

srp.Open " select * from customer_data", cn, adOpenDynamic, adLockOptimistic
If ccid.Text = "" Then
MsgBox ("Please Search any Record")
Else

spre = Trim(Right(ccid.Text, 11)) - 1

spre1 = rs!bid + spre


f = 0
    Do While srp.EOF <> True
    
    If spre1 = srp!Id Then
    f = 1
ccid.Text = srp!Id
ctype.Text = srp!actype
cifsc.Text = srp!IFSC
cac.Text = srp!ACNo
cname.Text = srp!Name
CFName.Text = srp!FName
CMName.Text = srp!MName
CDOB.Text = srp!DOB
CGender.Text = srp!Gender
CMobile.Text = srp!Mobile
CMail.Text = srp!Email
CLocation.Text = srp!Address
CPic.Picture = LoadPicture(srp!PImg)
CPictext.Text = srp!PImg
CSign.Picture = LoadPicture(srp!sImg)
CSigntext.Text = srp!sImg
SearchBox.Text = ccid.Text
If IsNull(srp!aadhar) Then
Caadh.Text = ""
MsgBox "KYC Pending, Update Aadhar Data", vbInformation
Caadh.Enabled = True
Caadh.SetFocus
Else
Caadh.Text = srp!aadhar
End If
Exit Do
End If
srp.MoveNext
Loop
If f = 0 Then
Snre1 = (rs!bid & (snre + 1))
If ccid.Text = Snre1 Then
MsgBox "This is First Record", vbInformation
Else
If ccid.Text <> Snre1 Then
MsgBox "You Can't See Record of other branch, but you can search individually", vbInformation
End If
End If
End If
End If
con.Close
rs.Close
End Sub

Private Sub Cross_Click()
Profile.Visible = False
Cross.Visible = False
LName.Visible = False
Pid.Visible = False
Pdoj.Visible = False
Pphone.Visible = False
Pmail.Visible = False
Paddress.Visible = False
PImg.Visible = False
FDPer.Visible = False

End Sub

Private Sub CSignedit_Click()
cname.Enabled = False
CFName.Enabled = False
CMName.Enabled = False
CDOB.Enabled = False
CGender.Enabled = False
CMobile.Enabled = False
CMail.Enabled = False
CLocation.Enabled = False
Caadh.Enabled = False
CD4.Filter = "JPG File | *.jpg|GIF File|*.gif"
CD4.ShowOpen
If CD4.FileName <> "" Then
CPic.Picture = LoadPicture(CD4.FileName)
CSigntext.Text = CD4.FileName
End If
End Sub

Private Sub CUpdate_Click()
Dim cdsearch As String
Dim cdsid As String
Dim cdnm As String
Dim cdfnm As String
Dim cdmnm As String
Dim cddob As String
Dim cdgdr As String
Dim cdmbl As String
Dim cdmail As String
Dim cdadrs As String
Dim cdpimg As String
Dim cdsimg As String
Dim cdaadhr As Integer

cdsearch = UCase(Trim(SearchBox.Text))
cdsid = UCase(Trim(ccid.Text))
cdnm = UCase(Trim(cname.Text))
cdfnm = UCase(Trim(CFName.Text))
cdmnm = UCase(Trim(CMName.Text))
cddob = UCase(Trim(CDOB.Text))
cdgdr = UCase(Trim(CGender.Text))
cdmbl = UCase(Trim(CMobile.Text))
cdmail = UCase(Trim(CMail.Text))
cdadrs = UCase(Trim(CLocation.Text))
cdpimg = Trim(CPictext.Text)
cdsimg = Trim(CSigntext.Text)
caadhr = Trim(Caadh.Text)

If ccid.Text = "" Then
MsgBox ("Search Customer First")
Else
If Len(CMobile.Text) > 10 Or Len(CMobile.Text) < 10 Then
      MsgBox "Enter the phone number in 10 digits!", vbExclamation, ""
      Cancel = True
      CMobile.SetFocus
   Else
Make_Connection
con.Open cn
con.Execute ("update customer_data set name='" & cdnm & "',fname='" & cdfnm & "',mname='" & cdmnm & "',gender='" & cdgdr & "',mobile='" & cdmbl & "',email='" & cdmail & "',address='" & cdadrs & "',simg='" & cdsimg & "',pimg='" & cdpimg & "',aadhar='" & caadhr & "' where id='" & cdsid & "' ")
con.Execute ("commit")
MsgBox ("A/C No. " + cac.Text + " Updated Successfully")

ccid.Text = ""
ctype.Text = ""
cifsc.Text = ""
cac.Text = ""
cname.Text = ""
CFName.Text = ""
CMName.Text = ""
CDOB.Text = ""
CGender.Text = ""
CMobile.Text = ""
CMail.Text = ""
CLocation.Text = ""
Caadh.Text = ""
CPic.Picture = Nothing
CSign.Picture = Nothing

cname.Enabled = False
CFName.Enabled = False
CMName.Enabled = False
CDOB.Enabled = False
CGender.Enabled = False
CMobile.Enabled = False
CMail.Enabled = False
CLocation.Enabled = False
Caadh.Enabled = False
con.Close
End If
End If
End Sub

Private Sub Current_Click()
FDPer.Visible = False
End Sub

Private Sub CustomerPage_Click()
SearchBox.ToolTipText = "Enter ID"
ccid.Text = ""
ctype.Text = ""
cifsc.Text = ""
cac.Text = ""
CPic.Picture = LoadPicture()
CSign.Picture = LoadPicture()
cname.Text = ""
CFName.Text = ""
CMName.Text = ""
CDOB.Text = ""
CGender.Text = ""
CMobile.Text = ""
CMail.Text = ""
CLocation.Text = ""
Caadh.Text = ""

Profile.Visible = False
Cross.Visible = False
LName.Visible = False
Pid.Visible = False
Pdoj.Visible = False
Pphone.Visible = False
Pmail.Visible = False
Paddress.Visible = False
PImg.Visible = False
ACancel.Visible = False
LOpen.Visible = False
FDPer.Visible = False

SearchTran.Visible = False
BankBal.Visible = False
Stdep.Visible = False
Stwid.Visible = False
Sttra.Visible = False
Sbrnm.Visible = False
Scod.Visible = False
Sifs.Visible = False
Sintr.Visible = False
Sloc.Visible = False
TransactionList.Visible = False

WithdrawBTN.Visible = False
DepositBtn.Visible = False

SCustom.Visible = False
Bal.Visible = False
wamt.Visible = False
Wbtn.Visible = False
StatementList.Visible = False

Saving.Visible = False
Current.Visible = False
FD.Visible = False
CID.Visible = False
OBal.Visible = False
OName.Visible = False
FName.Visible = False
MName.Visible = False
DOB.Visible = False
DTPicker1.Visible = False
Gender.Visible = False
Male.Visible = False
Female.Visible = False
Transgender.Visible = False
Mobile.Visible = False
Email.Visible = False
Address.Visible = False
PLogout.Visible = False
LogYes.Visible = False
LogNo.Visible = False
Hifsc.Visible = False
HomeCover.Visible = False
BCode.Visible = False
Hmbl.Visible = False
Hmail.Visible = False
Haddress.Visible = False
Twitter.Visible = False
Linkedin.Visible = False
facebook.Visible = False
Instagram.Visible = False
Website.Visible = False
Download1.Visible = False
Download2.Visible = False
APImg.Visible = False
UploadPImg.Visible = False
ASignImg.Visible = False
ASignUpload.Visible = False
APimgtext.Visible = False
ASignText.Visible = False

Enm.Visible = False
Enm1.Visible = False
Enm2.Visible = False
Enm3.Visible = False
Enm4.Visible = False
Enm5.Visible = False
Enm6.Visible = False
Enm7.Visible = False
TBD.Visible = False
TBD1.Visible = False
TBD2.Visible = False
TBD3.Visible = False

SEmpBtn.Visible = False
EmpNm.Visible = False
EmpPosition.Visible = False
EmpDOJ.Visible = False
EmpEID.Visible = False
EmpPhone.Visible = False
EmpMail.Visible = False
EmpSal.Visible = False
EmpAdd.Visible = False

SearchS.Visible = False
SObal.Visible = False
STDp.Visible = False
STWd.Visible = False
SCbal.Visible = False
SAcNo.Visible = False
SAtyp.Visible = False
SAIfsc.Visible = False
SIra.Visible = False
SNom.Visible = False

SearchBox.Visible = True
SearchCustomer.Visible = True
CPrev.Visible = True
CNext.Visible = True
CUpdate.Visible = True
CDel.Visible = True
CPic.Visible = True
CSign.Visible = True
CPicedit.Visible = True
CSignedit.Visible = True
ccid.Visible = True
ctype.Visible = True
cifsc.Visible = True
cac.Visible = True
cname.Visible = True
CNedit.Visible = True
CFName.Visible = True
CFedit.Visible = True
CMName.Visible = True
CMedit.Visible = True
CDOB.Visible = True
CDedit.Visible = True
CGender.Visible = True
CGedit.Visible = True
CMobile.Visible = True
CMoedit.Visible = True
CMail.Visible = True
CMaedit.Visible = True
CLocation.Visible = True
CLoedit.Visible = True
Caadh.Visible = True
Caadedit.Visible = True

Set Me.Picture = Me.CustomerImg
End Sub

Private Sub DepositBTN_Click()
Dim sfres As New ADODB.Recordset

Dim wduid As String
Dim dwduid As String

Dim wdifsc As String
Dim wdacno As String
Dim wdid As String
Dim wdname As String
Dim wdmobile As String
Dim wddate As String
Dim wddebit As String
Dim wdcredit As String
Dim wdbalance As String
Dim Vgbk As String
Dim tdep As String
Dim tmode As String
Dim vtdep1 As String
Dim dtypfi As String
dtypfi = "Credit"

Vgbk = ccid.Text
tmode = "Cash"

If ccid.Text = "" Then
MsgBox ("Please Search Customer")
SearchBox.SetFocus
Else
If wamt.Text = "" Then
MsgBox ("Please Enter Amount")
wamt.SetFocus
Else
Make_Connection
con.Open cn
sfres.Open " select * from customer_data where acno='" & Vgbk & "'", cn, adOpenDynamic, adLockOptimistic
banbal.Open " select * from admin_login", cn, adOpenDynamic, adLockOptimistic


wdifsc = sfres!IFSC
wdacno = sfres!ACNo
wdid = sfres!Id
wdname = sfres!Name
wdmobile = sfres!Mobile
wddate = Now()
wddebit = "0"
wdcredit = wamt.Text
wdbalance = Trim(sfres!Balance + wamt.Text)

If IsNull(banbal!bank_bal) Then
con.Execute ("update admin_login set bank_bal= 0")
con.Execute ("commit")
banbal.Close
banbal.Open " select bank_bal from admin_login", cn, adOpenDynamic, adLockOptimistic
bbala = Trim(banbal!bank_bal + wamt.Text)
Else
bbala = Trim(banbal!bank_bal + wamt.Text)
End If


wdtid = sfres!ltid + 1
tdep = Trim(sfres!tdeposit + wamt.Text)
con.Execute (" ALTER SESSION SET NLS_DATE_FORMAT='dd-mm-yyyy hh:mi:ss AM'")
con.Execute ("insert into customer_statement (ifsc,acno,id,name,mobile,dates,credit,debit,balance,tid,tmode,ttype)values('" & wdifsc & "','" & wdacno & "','" & wdid & "','" & wdname & "','" & wdmobile & "','" & wddate & "','" & wdcredit & "','" & wddebit & "','" & wdbalance & "','" & wdtid & "','" & tmode & "','" & dtypfi & "')")
con.Execute ("insert into Bank_Transaction (dates,tid,ttype,tmode,id,name,amount,acno,ifsc)values('" & wddate & "','" & wdtid & "','" & dtypfi & "','" & tmode & "','" & wdid & "','" & wdname & "','" & wdcredit & "','" & wdacno & "','" & wdifsc & "')")
con.Execute ("commit")
con.Execute ("update admin_login set bank_bal='" & bbala & "'")
con.Execute ("update customer_data set ltid='" & wdtid & "',balance='" & wdbalance & "',tdeposit='" & tdep & "' where id='" & wdid & "' ")
con.Execute ("commit")

vtdep1 = Trim(banbal!tdeposit + wamt.Text)
con.Execute ("update admin_login set tdeposit= '" & vtdep1 & "'")
con.Execute ("commit")
MsgBox ("Deposited Successfully, Your Balance is " + wdbalance)

Bal.Caption = wdbalance
wamt.Text = ""
banbal.Close
con.Close
End If
End If
End Sub

Private Sub DepositPage_Click()
Set Me.Picture = Me.DepositImg
SearchBox.ToolTipText = "Enter ID"
ccid.Text = ""
ctype.Text = ""
cifsc.Text = ""
cac.Text = ""
Bal.Caption = ""
CPic.Picture = LoadPicture()
CSign.Picture = LoadPicture()

Profile.Visible = False
Cross.Visible = False
LName.Visible = False
Pid.Visible = False
Pdoj.Visible = False
Pphone.Visible = False
Pmail.Visible = False
Paddress.Visible = False
PImg.Visible = False
ACancel.Visible = False
LOpen.Visible = False
WithdrawBTN.Visible = False
DepositBtn.Visible = True

SearchTran.Visible = False
BankBal.Visible = False
Stdep.Visible = False
Stwid.Visible = False
Sttra.Visible = False
Sbrnm.Visible = False
Scod.Visible = False
Sifs.Visible = False
Sintr.Visible = False
Sloc.Visible = False
TransactionList.Visible = False

SCustom.Visible = True
Bal.Visible = True
wamt.Visible = True
Wbtn.Visible = False
StatementList.Visible = False

Saving.Visible = False
Current.Visible = False
FD.Visible = False
CID.Visible = False
OBal.Visible = False
OName.Visible = False
FName.Visible = False
MName.Visible = False
DOB.Visible = False
DTPicker1.Visible = False
Gender.Visible = False
Male.Visible = False
Female.Visible = False
Transgender.Visible = False
Mobile.Visible = False
Email.Visible = False
Address.Visible = False
PLogout.Visible = False
LogYes.Visible = False
LogNo.Visible = False
Hifsc.Visible = False
HomeCover.Visible = False
BCode.Visible = False
Hmbl.Visible = False
Hmail.Visible = False
Haddress.Visible = False
Twitter.Visible = False
Linkedin.Visible = False
facebook.Visible = False
Instagram.Visible = False
Website.Visible = False
Download1.Visible = False
Download2.Visible = False
APImg.Visible = False
UploadPImg.Visible = False
ASignImg.Visible = False
ASignUpload.Visible = False
APimgtext.Visible = False
ASignText.Visible = False
FDPer.Visible = False

SearchS.Visible = False
SObal.Visible = False
STDp.Visible = False
STWd.Visible = False
SCbal.Visible = False
SAcNo.Visible = False
SAtyp.Visible = False
SAIfsc.Visible = False
SIra.Visible = False
SNom.Visible = False

SearchBox.Visible = True
SearchCustomer.Visible = False
CPrev.Visible = False
CNext.Visible = False
CUpdate.Visible = False
CDel.Visible = False
CPic.Visible = True
CSign.Visible = True
CPicedit.Visible = False
CSignedit.Visible = False
ccid.Visible = True
ctype.Visible = True
cifsc.Visible = True
cac.Visible = True
cname.Visible = False
CNedit.Visible = False
CFName.Visible = False
CFedit.Visible = False
CMName.Visible = False
CMedit.Visible = False
CDOB.Visible = False
CDedit.Visible = False
CGender.Visible = False
CGedit.Visible = False
CMobile.Visible = False
CMoedit.Visible = False
CMail.Visible = False
CMaedit.Visible = False
CLocation.Visible = False
CLoedit.Visible = False
Caadh.Visible = False
Caadedit.Visible = False

Enm.Visible = False
Enm1.Visible = False
Enm2.Visible = False
Enm3.Visible = False
Enm4.Visible = False
Enm5.Visible = False
Enm6.Visible = False
Enm7.Visible = False
TBD.Visible = False
TBD1.Visible = False
TBD2.Visible = False
TBD3.Visible = False

SEmpBtn.Visible = False
EmpNm.Visible = False
EmpPosition.Visible = False
EmpDOJ.Visible = False
EmpEID.Visible = False
EmpPhone.Visible = False
EmpMail.Visible = False
EmpSal.Visible = False
EmpAdd.Visible = False

End Sub
Private Sub DOB_Click()
DOB.Text = DTPicker1.Value
End Sub

Private Sub DOB_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
DOB_Click
End If
End Sub

Private Sub Download1_Click()
Shell ("explorer https://www.sbi.co.in/documents/16012/1557541/121120-Account+Opening+Form+for+Individuals.pdf")
End Sub

Private Sub Download2_Click()
Shell ("explorer https://www.bbkindia.com/pdf/form60.pdf")
End Sub

Private Sub DTPicker1_Click()
DOB.Text = DTPicker1.Value
End Sub
Private Sub EmpAdd_Change()
If EmpAdd.Locked = False Then
EmpAdd.FontBold = True
End If
End Sub

Private Sub EmpDel_Click()
Dim empids As String
empids = EmpEID.Text
If EmpEID.Text = "" Then
MsgBox ("Search Employee First")
Else
answer = MsgBox("Are you sure you want to close account of " & vbCrLf & "'" & cname.Text & "'", vbCritical + vbYesNo, "Warning")
If answer = vbYes Then
Make_Connection
con.Open cn
con.Execute ("delete from admin_login where id='" & empids & "' ")
con.Execute ("commit")
MsgBox ("Employee ID " + EmpEID.Text + " Shifted to Alumni")

EmpNm.Text = ""
EmpPosition.Text = ""
EmpDOJ.Text = ""
EmpEID.Text = ""
EmpPhone.Text = ""
EmpMail.Text = ""
EmpSal.Text = ""
EmpAdd.Text = ""
SearchBox.SetFocus

EmpNm.Locked = True
EmpPosition.Locked = True
EmpDOJ.Locked = True
EmpEID.Locked = True
EmpPhone.Locked = True
EmpMail.Locked = True
EmpSal.Locked = True
EmpAdd.Locked = True
con.Close
EmployeePage_Click
End If
End If
End Sub

Private Sub EmpEdit_Click()
If EmpEID.Text = "" Then
MsgBox ("Search Employee First")
Else
EmpNm.Locked = False
EmpNm.SetFocus
EmpPosition.Locked = False

EmpPhone.Locked = False
EmpMail.Locked = False
EmpSal.Locked = False
EmpAdd.Locked = False
End If
End Sub

Private Sub EmployeePage_Click()
Set Me.Picture = Me.EmployeeImg
SearchBox.ToolTipText = "Enter Employee ID"
Profile.Visible = False
Cross.Visible = False
LName.Visible = False
Pid.Visible = False
Pdoj.Visible = False
Pphone.Visible = False
Pmail.Visible = False
Paddress.Visible = False
PImg.Visible = False
ACancel.Visible = False
LOpen.Visible = False
FDPer.Visible = False

SearchTran.Visible = False
BankBal.Visible = False
Stdep.Visible = False
Stwid.Visible = False
Sttra.Visible = False
Sbrnm.Visible = False
Scod.Visible = False
Sifs.Visible = False
Sintr.Visible = False
Sloc.Visible = False
TransactionList.Visible = False

WithdrawBTN.Visible = False
DepositBtn.Visible = False

SCustom.Visible = False
Bal.Visible = False
wamt.Visible = False
Wbtn.Visible = False
StatementList.Visible = False

Saving.Visible = False
Current.Visible = False
FD.Visible = False
CID.Visible = False
OBal.Visible = False
OName.Visible = False
FName.Visible = False
MName.Visible = False
DOB.Visible = False
DTPicker1.Visible = False
Gender.Visible = False
Male.Visible = False
Female.Visible = False
Transgender.Visible = False
Mobile.Visible = False
Email.Visible = False
Address.Visible = False
PLogout.Visible = False
LogYes.Visible = False
LogNo.Visible = False
Hifsc.Visible = False
HomeCover.Visible = False
BCode.Visible = False
Hmbl.Visible = False
Hmail.Visible = False
Haddress.Visible = False
Twitter.Visible = False
Linkedin.Visible = False
facebook.Visible = False
Instagram.Visible = False
Website.Visible = False
Download1.Visible = False
Download2.Visible = False
APImg.Visible = False
UploadPImg.Visible = False
ASignImg.Visible = False
ASignUpload.Visible = False
APimgtext.Visible = False
ASignText.Visible = False

SearchBox.Visible = True
SearchCustomer.Visible = False
CPrev.Visible = False
CNext.Visible = False
CUpdate.Visible = False
CDel.Visible = False
CPic.Visible = False
CSign.Visible = False
CPicedit.Visible = False
CSignedit.Visible = False
ccid.Visible = False
ctype.Visible = False
cifsc.Visible = False
cac.Visible = False
cname.Visible = False
CNedit.Visible = False
CFName.Visible = False
CFedit.Visible = False
CMName.Visible = False
CMedit.Visible = False
CDOB.Visible = False
CDedit.Visible = False
CGender.Visible = False
CGedit.Visible = False
CMobile.Visible = False
CMoedit.Visible = False
CMail.Visible = False
CMaedit.Visible = False
CLocation.Visible = False
CLoedit.Visible = False
Caadh.Visible = False
Caadedit.Visible = False

SearchS.Visible = False
SObal.Visible = False
STDp.Visible = False
STWd.Visible = False
SCbal.Visible = False
SAcNo.Visible = False
SAtyp.Visible = False
SAIfsc.Visible = False
SIra.Visible = False
SNom.Visible = False

Enm.Visible = True
Enm1.Visible = True
Enm2.Visible = True
Enm3.Visible = True
Enm4.Visible = True
Enm5.Visible = True
Enm6.Visible = True
Enm7.Visible = True
TBD.Visible = True
TBD1.Visible = True
TBD2.Visible = True
TBD3.Visible = True

SEmpBtn.Visible = True
EmpNm.Visible = True
EmpPosition.Visible = True
EmpDOJ.Visible = True
EmpEID.Visible = True
EmpPhone.Visible = True
EmpMail.Visible = True
EmpSal.Visible = True
EmpAdd.Visible = True

EmpNew.Visible = True
EmpEdit.Visible = True
EmpDel.Visible = True
EmpSave.Visible = True

Dim esrp As New ADODB.Recordset
Dim esrp1 As New ADODB.Recordset
Dim esrp2 As New ADODB.Recordset
Dim esrp3 As New ADODB.Recordset
Dim esrp4 As New ADODB.Recordset
Dim esrp5 As New ADODB.Recordset
Dim esrp6 As New ADODB.Recordset
Dim esrp7 As New ADODB.Recordset

Dim edsrp As New ADODB.Recordset
Dim edsrp1 As New ADODB.Recordset
Dim edsrp2 As New ADODB.Recordset

Dim evi As String
Dim evi1 As String
Dim evi2 As String
Dim evi3 As String

Dim nu As Integer
Dim nu1 As Integer
Dim nu2 As Integer
Dim nu3 As Integer
Dim nu4 As Integer
Dim nu5 As Integer
Dim nu6 As Integer
Dim nu7 As Integer

EmpNm.Text = ""
EmpPosition.Text = ""
EmpDOJ.Text = ""
EmpEID.Text = ""
EmpPhone.Text = ""
EmpMail.Text = ""
EmpSal.Text = ""
EmpAdd.Text = ""

Make_Connection
con.Open cn
Enm.Clear
Enm1.Clear
Enm2.Clear
Enm3.Clear
Enm4.Clear
Enm5.Clear
Enm6.Clear
Enm7.Clear

esrp.CursorLocation = adUseClient
edsrp.CursorLocation = adUseClient
edsrp1.CursorLocation = adUseClient
edsrp2.CursorLocation = adUseClient

esrp.Open "SELECT name From admin_login", cn, adOpenKeyset, adLockOptimistic, adCmdText
esrp1.Open "SELECT post From admin_login", cn, adOpenDynamic, adLockUnspecified
esrp2.Open "SELECT doj From admin_login", cn, adOpenDynamic, adLockUnspecified
esrp3.Open "SELECT id From admin_login", cn, adOpenDynamic, adLockUnspecified
esrp4.Open "SELECT phone From admin_login", cn, adOpenDynamic, adLockUnspecified
esrp5.Open "SELECT email From admin_login", cn, adOpenDynamic, adLockUnspecified
esrp6.Open "SELECT salary From admin_login", cn, adOpenDynamic, adLockUnspecified
esrp7.Open "SELECT address From admin_login", cn, adOpenDynamic, adLockUnspecified

edsrp.Open "SELECT distinct branch_pincode from admin_login", cn, adOpenDynamic, adLockUnspecified
evi1 = edsrp.RecordCount
edsrp1.Open "SELECT id from customer_data", cn, adOpenDynamic, adLockUnspecified
evi2 = edsrp1.RecordCount
edsrp2.Open "SELECT distinct branch_pincode from admin_login", cn, adOpenDynamic, adLockUnspecified
evi3 = edsrp2.RecordCount


evi = esrp.RecordCount
esrp.MoveFirst
esrp1.MoveFirst
esrp2.MoveFirst
esrp3.MoveFirst
esrp4.MoveFirst
esrp5.MoveFirst
esrp6.MoveFirst
esrp7.MoveFirst

TBD.Caption = evi1
TBD1.Caption = evi
TBD2.Caption = evi2
TBD3.Caption = evi3

nu = 1
Do While nu <= evi
Enm.AddItem esrp.GetString(adClipString, 1)
Enm.AddItem vbNewLine
nu = nu + 1
Loop

nu1 = 1
Do While nu1 <= evi

Enm1.AddItem esrp1.GetString(adClipString, 1)
Enm1.AddItem vbNewLine
nu1 = nu1 + 1
Loop

nu2 = 1
Do While nu2 <= evi

Enm2.AddItem esrp2.GetString(adClipString, 1)
Enm2.AddItem vbNewLine
nu2 = nu2 + 1
Loop

nu3 = 1
Do While nu3 <= evi

Enm3.AddItem esrp3.GetString(adClipString, 1)
Enm3.AddItem vbNewLine
nu3 = nu3 + 1
Loop

nu4 = 1
Do While nu4 <= evi

Enm4.AddItem esrp4.GetString(adClipString, 1)
Enm4.AddItem vbNewLine
nu4 = nu4 + 1
Loop

nu5 = 1
Do While nu5 <= evi

Enm5.AddItem esrp5.GetString(adClipString, 1)
Enm5.AddItem vbNewLine
nu5 = nu5 + 1
Loop

nu6 = 1
Do While nu6 <= evi

Enm6.AddItem esrp6.GetString(adClipString, 1)
Enm6.AddItem vbNewLine
nu6 = nu6 + 1
Loop

nu7 = 1
Do While nu7 <= evi

Enm7.AddItem esrp7.GetString(adClipString, 1)
Enm7.AddItem vbNewLine
nu7 = nu7 + 1
Loop

con.Close
End Sub

Private Sub EmpMail_Change()
If EmpMail.Locked = False Then
EmpMail.FontBold = True
End If
End Sub

Private Sub EmpNm_Click()
If EmpNm.Locked = False Then
EmpNm.FontBold = True
End If
End Sub

Private Sub EmpPhone_Change()
If EmpPhone.Locked = False Then
EmpPhone.FontBold = True
End If
End Sub

Private Sub EmpPhone_KeyPress(KeyAscii As Integer)
If KeyAscii <> 48 And KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 52 And KeyAscii <> 53 And KeyAscii <> 54 And KeyAscii <> 55 And KeyAscii <> 56 And KeyAscii <> 57 And KeyAscii <> 8 And KeyAscii <> 13 Then
    KeyAscii = 0
    MsgBox ("Please Enter Numbers Only")
    End If
End Sub
Private Sub EmpSal_Change()
If EmpSal.Locked = False Then
EmpSal.FontBold = True
End If
End Sub

Private Sub EmpSal_KeyPress(KeyAscii As Integer)
If KeyAscii <> 48 And KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 52 And KeyAscii <> 53 And KeyAscii <> 54 And KeyAscii <> 55 And KeyAscii <> 56 And KeyAscii <> 57 And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 46 Then
    KeyAscii = 0
    MsgBox ("Please Enter Numbers Only")
    End If
End Sub

Private Sub EmpSave_Click()
Dim empsnm As String
Dim empspos As String
Dim empsdoj As String
Dim empseid As String
Dim empsphn As String
Dim empsmail As String
Dim empssal As String
Dim empsadr As String

empsnm = UCase(Trim(EmpNm.Text))
empspos = UCase(Trim(EmpPosition.Text))
empsdoj = UCase(Trim(EmpDOJ.Text))
empseid = UCase(Trim(EmpEID.Text))
empsphn = UCase(Trim(EmpPhone.Text))
empsmail = UCase(Trim(EmpMail.Text))
empssal = UCase(Trim(EmpSal.Text))
empsadr = UCase(Trim(EmpAdd.Text))

If EmpEID.Text = "" Then
MsgBox ("Search Employee First")
Else

Make_Connection
con.Open cn
If Len(EmpPhone.Text) > 10 Or Len(EmpPhone.Text) < 10 Then
      MsgBox "Enter the phone number in 10 digits!", vbExclamation, ""
      Cancel = True
      EmpPhone.SetFocus
   Else
con.Execute ("update admin_login set name='" & empsnm & "',post='" & empspos & "',phone='" & empsphn & "',email='" & empsmail & "',salary='" & empssal & "',address='" & empsadr & "' where id='" & empseid & "' ")
con.Execute ("commit")
MsgBox ("Employee Id " + EmpEID.Text + " Updated Successfully")

EmpNm.Text = ""
EmpPosition.Text = ""
EmpDOJ.Text = ""
EmpEID.Text = ""
EmpPhone.Text = ""
EmpMail.Text = ""
EmpSal.Text = ""
EmpAdd.Text = ""
SearchBox.SetFocus

EmpNm.Locked = True
EmpPosition.Locked = True
EmpDOJ.Locked = True
EmpEID.Locked = True
EmpPhone.Locked = True
EmpMail.Locked = True
EmpSal.Locked = True
EmpAdd.Locked = True
End If
con.Close
End If
End Sub
Private Sub Form_Load()
PlaySound App.Path & "\Welcome_To_SA_Bank.wav"
Dim kk As String
Dim st1 As String
HomeCover.Visible = True
Hifsc.Visible = True
BCode.Visible = True
Hmbl.Visible = True
Hmail.Visible = True
Haddress.Visible = True
Twitter.Visible = True
Linkedin.Visible = True
facebook.Visible = True
Instagram.Visible = True
Website.Visible = True
Download1.Visible = True
Download2.Visible = True

Make_Connection
con.Open cn
If Login.Combo1.Text = "Admin" Then
rs.Open " select * from admin_login", cn, adOpenDynamic, adLockOptimistic
kk = LCase(Trim(Login.Id.Text))
    rs.MoveFirst
    f = 0
    Do While rs.EOF <> True
    st1 = LCase(Trim(rs!Id))
    If kk = st1 Then
    f = 1
Label1.Caption = "Welcome " + StrConv(rs!Name, vbProperCase)
LName.Caption = rs!Name
Label5.Caption = rs!Branch_name
Pid.Caption = rs!Id
Pdoj.Caption = rs!doj
Pphone.Caption = rs!phone
Pmail.Caption = rs!Email
Paddress.Caption = rs!Address
Hifsc.Caption = rs!IFSC
BCode.Caption = rs!bid
Hmbl.Caption = rs!phone
Hmail.Caption = rs!Email
Haddress.Caption = rs!Address
Twitter.Caption = rs!Twitter
Linkedin.Caption = rs!Linkedin
facebook.Caption = rs!facebook
Instagram.Caption = rs!Instagram
Website.Caption = rs!Website

Exit Do
End If
rs.MoveNext
Loop
If f = 0 Then
Label1.Caption = "No Record"

End If
rs.Close

    Else
    
If Login.Combo1.Text = "Customer" Then
zo.Open " select * from customer_data", cn, adOpenDynamic, adLockOptimistic
kk = LCase(Trim(Login.Id.Text))
    zo.MoveFirst
    f = 0
    Do While rs.EOF <> True
    st1 = LCase(Trim(rs!Id))
    If kk = st1 Then
    f = 1
Label1.Caption = "Welcome " + zo!Name
LName.Caption = rs!Name
Label5.Caption = zo!Branch_name
Pid.Caption = zo!Id
Pdoj.Caption = zo!doj
Pphone.Caption = zo!phone
Pmail.Caption = zo!Email
Paddress.Caption = zo!Address
Hifsc.Caption = zo!IFSC
Exit Do
End If
zo.MoveNext
Loop
If f = 0 Then
Label1.Caption = "No Record"

End If
rs.Close
End If
End If

con.Close
End Sub

Private Sub Date_Timer()
Date_Time.Caption = Now
End Sub

Private Sub HomePage_Click()
Profile.Visible = False
Cross.Visible = False
LName.Visible = False
Pid.Visible = False
Pdoj.Visible = False
Pphone.Visible = False
Pmail.Visible = False
Paddress.Visible = False
PImg.Visible = False
ACancel.Visible = False
LOpen.Visible = False
FDPer.Visible = False

SCustom.Visible = False
Bal.Visible = False
wamt.Visible = False
Wbtn.Visible = False

Saving.Visible = False
Current.Visible = False
FD.Visible = False
CID.Visible = False
OBal.Visible = False
OName.Visible = False
FName.Visible = False
MName.Visible = False
DOB.Visible = False
DTPicker1.Visible = False
Gender.Visible = False
Male.Visible = False
Female.Visible = False
Transgender.Visible = False
Mobile.Visible = False
Email.Visible = False
Address.Visible = False
PLogout.Visible = False
LogYes.Visible = False
LogNo.Visible = False
APImg.Visible = False
UploadPImg.Visible = False
ASignImg.Visible = False
ASignUpload.Visible = False
APimgtext.Visible = False
ASignText.Visible = False

SearchBox.Visible = False
SearchCustomer.Visible = False
CPrev.Visible = False
CNext.Visible = False
CUpdate.Visible = False
CDel.Visible = False
CPic.Visible = False
CSign.Visible = False
CPicedit.Visible = False
CSignedit.Visible = False
ccid.Visible = False
ctype.Visible = False
cifsc.Visible = False
cac.Visible = False
cname.Visible = False
CNedit.Visible = False
CFName.Visible = False
CFedit.Visible = False
CMName.Visible = False
CMedit.Visible = False
CDOB.Visible = False
CDedit.Visible = False
CGender.Visible = False
CGedit.Visible = False
CMobile.Visible = False
CMoedit.Visible = False
CMail.Visible = False
CMaedit.Visible = False
CLocation.Visible = False
CLoedit.Visible = False
Caadh.Visible = False
Caadedit.Visible = False

WithdrawBTN.Visible = False
DepositBtn.Visible = False
StatementList.Visible = False

SearchTran.Visible = False
BankBal.Visible = False
Stdep.Visible = False
Stwid.Visible = False
Sttra.Visible = False
Sbrnm.Visible = False
Scod.Visible = False
Sifs.Visible = False
Sintr.Visible = False
Sloc.Visible = False
TransactionList.Visible = False

Enm.Visible = False
Enm1.Visible = False
Enm2.Visible = False
Enm3.Visible = False
Enm4.Visible = False
Enm5.Visible = False
Enm6.Visible = False
Enm7.Visible = False
TBD.Visible = False
TBD1.Visible = False
TBD2.Visible = False
TBD3.Visible = False

SEmpBtn.Visible = False
EmpNm.Visible = False
EmpPosition.Visible = False
EmpDOJ.Visible = False
EmpEID.Visible = False
EmpPhone.Visible = False
EmpMail.Visible = False
EmpSal.Visible = False
EmpAdd.Visible = False

SearchS.Visible = False
SObal.Visible = False
STDp.Visible = False
STWd.Visible = False
SCbal.Visible = False
SAcNo.Visible = False
SAtyp.Visible = False
SAIfsc.Visible = False
SIra.Visible = False
SNom.Visible = False

HomeCover.Visible = True
Hifsc.Visible = True
BCode.Visible = True
Hmbl.Visible = True
Hmail.Visible = True
Haddress.Visible = True
Twitter.Visible = True
Linkedin.Visible = True
facebook.Visible = True
Instagram.Visible = True
Website.Visible = True
Download1.Visible = True
Download2.Visible = True

Set Me.Picture = Me.HomeImg
End Sub

Private Sub LogNo_Click()
PLogout.Visible = False
HomeCover.Visible = True
Hifsc.Visible = True
BCode.Visible = True
Hmbl.Visible = True
Hmail.Visible = True
Haddress.Visible = True
Twitter.Visible = True
Linkedin.Visible = True
facebook.Visible = True
Instagram.Visible = True
Website.Visible = True
Download1.Visible = True
Download2.Visible = True
Set Me.Picture = Me.HomeImg
End Sub

Private Sub LogoutPage_Click()
Profile.Visible = False
Cross.Visible = False
LName.Visible = False
Pid.Visible = False
Pdoj.Visible = False
Pphone.Visible = False
Pmail.Visible = False
Paddress.Visible = False
PImg.Visible = False
ACancel.Visible = False
LOpen.Visible = False
FDPer.Visible = False

WithdrawBTN.Visible = False
DepositBtn.Visible = False
StatementList.Visible = False

SearchTran.Visible = False
BankBal.Visible = False
Stdep.Visible = False
Stwid.Visible = False
Sttra.Visible = False
Sbrnm.Visible = False
Scod.Visible = False
Sifs.Visible = False
Sintr.Visible = False
Sloc.Visible = False
TransactionList.Visible = False

SCustom.Visible = False
Bal.Visible = False
wamt.Visible = False
Wbtn.Visible = False

Saving.Visible = False
Current.Visible = False
FD.Visible = False
CID.Visible = False
OBal.Visible = False
OName.Visible = False
FName.Visible = False
MName.Visible = False
DOB.Visible = False
DTPicker1.Visible = False
Gender.Visible = False
Male.Visible = False
Female.Visible = False
Transgender.Visible = False
Mobile.Visible = False
Email.Visible = False
Address.Visible = False
PLogout.Visible = True
LogYes.Visible = True
LogNo.Visible = True
Hifsc.Visible = False
HomeCover.Visible = False
BCode.Visible = False
Hmbl.Visible = False
Hmail.Visible = False
Haddress.Visible = False
Twitter.Visible = False
Linkedin.Visible = False
facebook.Visible = False
Instagram.Visible = False
Website.Visible = False
Download1.Visible = False
Download2.Visible = False
APImg.Visible = False
UploadPImg.Visible = False
ASignImg.Visible = False
ASignUpload.Visible = False
APimgtext.Visible = False
ASignText.Visible = False

SearchBox.Visible = False
SearchCustomer.Visible = False
CPrev.Visible = False
CNext.Visible = False
CUpdate.Visible = False
CDel.Visible = False
CPic.Visible = False
CSign.Visible = False
CPicedit.Visible = False
CSignedit.Visible = False
ccid.Visible = False
ctype.Visible = False
cifsc.Visible = False
cac.Visible = False
cname.Visible = False
CNedit.Visible = False
CFName.Visible = False
CFedit.Visible = False
CMName.Visible = False
CMedit.Visible = False
CDOB.Visible = False
CDedit.Visible = False
CGender.Visible = False
CGedit.Visible = False
CMobile.Visible = False
CMoedit.Visible = False
CMail.Visible = False
CMaedit.Visible = False
CLocation.Visible = False
CLoedit.Visible = False
Caadh.Visible = False
Caadedit.Visible = False

Enm.Visible = False
Enm1.Visible = False
Enm2.Visible = False
Enm3.Visible = False
Enm4.Visible = False
Enm5.Visible = False
Enm6.Visible = False
Enm7.Visible = False
TBD.Visible = False
TBD1.Visible = False
TBD2.Visible = False
TBD3.Visible = False

SEmpBtn.Visible = False
EmpNm.Visible = False
EmpPosition.Visible = False
EmpDOJ.Visible = False
EmpEID.Visible = False
EmpPhone.Visible = False
EmpMail.Visible = False
EmpSal.Visible = False
EmpAdd.Visible = False

SearchS.Visible = False
SObal.Visible = False
STDp.Visible = False
STWd.Visible = False
SCbal.Visible = False
SAcNo.Visible = False
SAtyp.Visible = False
SAIfsc.Visible = False
SIra.Visible = False
SNom.Visible = False

End Sub

Private Sub LogYes_Click()
Unload Me
Login.Show
End Sub

Private Sub LOpen_Click()
Dim dtype As String
Dim did As String
Dim dnm As String
Dim dfnm As String
Dim dmnm As String
Dim ddob As String
Dim dgdr As String
Dim dmbl As String
Dim dmail As String
Dim dadrs As String
Dim dbranch As String
Dim difsc As String
Dim dacn As String
Dim dobal As String
Dim dpimg As String
Dim dsimg As String
Dim dltid As String
Dim aodat As String
Dim adeb As String
Dim intr As String
Dim fdpd As String
Dim nyhc As String
Dim vtdep As String
Dim otypfi As String
otypfi = "Credit"


If Saving.Value = True Then
dtype = Saving.Caption
FDPer.Text = " "
Else
If Current.Value = True Then
dtype = Current.Caption
FDPer.Text = " "
Else
If FD.Value = True Then
dtype = FD.Caption
FDPer.Visible = True
End If
End If
End If
dacn = Trim(ACNo.Text)
did = UCase(Trim(CID.Text))
dobal = Trim(OBal.Text)
dnm = UCase(Trim(OName.Text))
dfnm = UCase(Trim(FName.Text))
dmnm = UCase(Trim(MName.Text))
ddob = Trim(DOB.Text)
dpimg = Trim(APimgtext.Text)
dsimg = Trim(ASignText.Text)
dltid = Trim(dacn + "1")
adeb = "0"
aodat = Trim(Now())
fdpd = Trim(FDPer.Text)
nyhc = "OBal, Cash"

If Male.Value = True Then
dgdr = Male.Caption
Else
If Female.Value = True Then
dgdr = Female.Caption
Else
If Transgender.Value = True Then
dgdr = Transgender.Caption
End If
End If
End If

If Saving.Value = True Then
intr = "4.00%"
Else
If Current.Value = True Then
intr = "2.50%"
Else
If FD.Value = True Then
intr = "6.00%"
End If
End If
End If

dmbl = (Trim(Mobile.Text))
dmail = UCase(Trim(Email.Text))
dadrs = UCase(Trim(Address.Text))
dbranch = UCase(Trim(Label5.Caption))
difsc = UCase(Trim(Hifsc.Caption))

Make_Connection
con.Open cn
If Saving.Value = False Xor Current.Value = False Xor FD.Value = False Then
MsgBox ("Please Select Account Type")
Else

If dtype = FD.Caption And FDPer.Text = "" Then
MsgBox "Choose FD Period", vbQuestion
FDPer.Visible = True
Else
If dtype = FD.Caption And FDPer.Text = " " Then
MsgBox "Choose FD Period"
Else
If CID.Text = "" Then
MsgBox ("Customer Id Can't be Empty, Please click on Open Account")
Else
If OBal.Text = "" Then
MsgBox ("Please Enter Opening Balance")
OBal.SetFocus
Else
If ACNo.Text = "" Then
MsgBox ("Please click on Open Account")
Else
If OName.Text = "" Then
MsgBox ("Please Enter Name")
OName.SetFocus
Else
If FName.Text = "" Then
MsgBox ("Please Enter Father's Name")
FName.SetFocus
Else
If MName.Text = "" Then
MsgBox ("Please Enter Mother's Name")
MName.SetFocus
Else
If DOB.Text = "" Then
MsgBox ("Please Enter Date of Birth in DD-MM-YYYY format")
DOB.SetFocus
Else
If Male.Value = False Xor Female.Value = False Xor Transgender.Value = False Then
MsgBox ("Please Select Gender")
Else
If Len(Mobile.Text) > 10 Or Len(Mobile.Text) < 10 Then
      MsgBox "Enter Mobile Number in 10 digits!", vbExclamation, ""
      Cancel = True
    Mobile.SetFocus
Else
If Email.Text = "" Then
MsgBox ("Please Enter Valid E-mail Id")
Email.SetFocus
Else
If Address.Text = "" Then
MsgBox ("Please Enter Address")
Address.SetFocus
Else
If APImg.Picture = Empty Then
MsgBox ("Please Upload Your Photo")

Else
If ASignImg.Picture = Empty Then
MsgBox ("Please Upload Your Signature")

Else
con.Execute ("ALTER SESSION SET nls_date_format='dd-mm-yyyy'")

con.Execute ("insert into customer_data (branch_name,ifsc,acno,actype,id,obalance,name,fname,mname,dob,gender,mobile,email,address,simg,pimg,balance,ltid,tdeposit,twithdrawl,aodate,interestrate,fdperiod)values('" & dbranch & "','" & difsc & "','" & dacn & "','" & dtype & "','" & did & "','" & dobal & "','" & dnm & "','" & dfnm & "','" & dmnm & "','" & ddob & "','" & dgdr & "','" & dmbl & "','" & dmail & "','" & dadrs & "','" & dsimg & "','" & dpimg & "','" & dobal & "','" & dltid & "','" & dobal & "','" & adeb & "','" & aodat & "','" & intr & "','" & fdpd & "')")
con.Execute ("commit")

con.Execute (" ALTER SESSION SET NLS_DATE_FORMAT='dd-mm-yyyy hh:mi:ss AM'")
con.Execute ("insert into customer_statement (ifsc,acno,id,name,mobile,dates,credit,debit,balance,tid,tmode,ttype)values('" & difsc & "','" & dacn & "','" & did & "','" & dnm & "','" & dmbl & "','" & aodat & "','" & dobal & "','" & adeb & "','" & dobal & "','" & dltid & "','" & nyhc & "','" & otypfi & "')")
con.Execute ("insert into Bank_Transaction (dates,tid,ttype,tmode,id,name,amount,acno,ifsc)values('" & aodat & "','" & dltid & "','" & otypfi & "','" & nyhc & "','" & did & "','" & dnm & "','" & dobal & "','" & dacn & "','" & difsc & "')")
con.Execute ("commit")

'Add Total Bank Balance Code Start
banbal.Open " select * from admin_login", cn, adOpenDynamic, adLockOptimistic
If IsNull(banbal!bank_bal) Then
con.Execute ("update admin_login set bank_bal= 0")
con.Execute ("commit")
banbal.Close
banbal.Open " select * from admin_login", cn, adOpenDynamic, adLockOptimistic
bbala = Trim(banbal!bank_bal + OBal.Text)
Else
bbala = Trim(banbal!bank_bal + OBal.Text)
End If
con.Execute ("update admin_login set bank_bal='" & bbala & "'")
con.Execute ("commit")
banbal.Close

banbal.Open
If IsNull(banbal!tdeposit) Then
con.Execute ("update admin_login set tdeposit= 0")
con.Execute ("commit")
banbal.Close
banbal.Open " select * from admin_login", cn, adOpenDynamic, adLockOptimistic
vtdep = Trim(banbal!tdeposit + OBal.Text)
con.Execute ("update admin_login set tdeposit= '" & vtdep & "'")
con.Execute ("commit")
Else
vtdep = Trim(banbal!tdeposit + OBal.Text)
con.Execute ("update admin_login set tdeposit= '" & vtdep & "'")
con.Execute ("commit")
End If
banbal.Close
'Code End


MsgBox ("Account Open Successfully, Your A/c No. is " + ACNo.Text)

ACNo.Text = ""
CID.Text = ""
OBal.Text = ""
OName.Text = ""
FName.Text = ""
MName.Text = ""
DOB.Text = ""
Mobile.Text = ""
Email.Text = ""
Address.Text = ""
APImg.Picture = Nothing
ASignImg.Picture = Nothing
OName.SetFocus
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
con.Close

End Sub

Private Sub Mobile_KeyPress(KeyAscii As Integer)
If KeyAscii <> 48 And KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 52 And KeyAscii <> 53 And KeyAscii <> 54 And KeyAscii <> 55 And KeyAscii <> 56 And KeyAscii <> 57 And KeyAscii <> 8 And KeyAscii <> 13 Then
    KeyAscii = 0
    MsgBox ("Please Enter Numbers Only")
    End If
End Sub

Private Sub OBal_Click()
If FD.Value = True Then
FDPer.Visible = True
Else
FDPer.Visible = False
End If
End Sub

Private Sub OBal_KeyPress(KeyAscii As Integer)
If KeyAscii <> 48 And KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 52 And KeyAscii <> 53 And KeyAscii <> 54 And KeyAscii <> 55 And KeyAscii <> 56 And KeyAscii <> 57 And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 46 Then
    KeyAscii = 0
    MsgBox ("Please Enter Numbers Only")
    End If
End Sub

Private Sub OpenAccountPage_Click()
Set Me.Picture = Me.AccountImg
Dim kt As String
Dim ju As String
Dim acn As String
Dim zc As New ADODB.Recordset

Saving.Value = True
Profile.Visible = False
Cross.Visible = False
LName.Visible = False
Pid.Visible = False
Pdoj.Visible = False
Pphone.Visible = False
Pmail.Visible = False
Paddress.Visible = False
PImg.Visible = False

SearchTran.Visible = False
BankBal.Visible = False
Stdep.Visible = False
Stwid.Visible = False
Sttra.Visible = False
Sbrnm.Visible = False
Scod.Visible = False
Sifs.Visible = False
Sintr.Visible = False
Sloc.Visible = False
TransactionList.Visible = False

SCustom.Visible = False
Bal.Visible = False
wamt.Visible = False
Wbtn.Visible = False

WithdrawBTN.Visible = False
DepositBtn.Visible = False

SearchS.Visible = False
SObal.Visible = False
STDp.Visible = False
STWd.Visible = False
SCbal.Visible = False
SAcNo.Visible = False
SAtyp.Visible = False
SAIfsc.Visible = False
SIra.Visible = False
SNom.Visible = False

Saving.Visible = True
Current.Visible = True
FD.Visible = True
CID.Visible = True
OBal.Visible = True
OName.Visible = True
FName.Visible = True
MName.Visible = True
DOB.Visible = True
DTPicker1.Visible = True
Gender.Visible = True
Male.Visible = True
Female.Visible = True
Transgender.Visible = True
Mobile.Visible = True
Email.Visible = True
Address.Visible = True
APImg.Visible = True
UploadPImg.Visible = True
ASignImg.Visible = True
ASignUpload.Visible = True
APimgtext.Visible = True
ASignText.Visible = True
ACancel.Visible = True
LOpen.Visible = True

PLogout.Visible = False
LogYes.Visible = False
LogNo.Visible = False
Hifsc.Visible = False
HomeCover.Visible = False
BCode.Visible = False
Hmbl.Visible = False
Hmail.Visible = False
Haddress.Visible = False
Twitter.Visible = False
Linkedin.Visible = False
facebook.Visible = False
Instagram.Visible = False
Website.Visible = False
Download1.Visible = False
Download2.Visible = False
SearchBox.Visible = False
SearchCustomer.Visible = False
CPrev.Visible = False
CNext.Visible = False
CUpdate.Visible = False
CDel.Visible = False
CPic.Visible = False
CSign.Visible = False
CPicedit.Visible = False
CSignedit.Visible = False
ccid.Visible = False
ctype.Visible = False
cifsc.Visible = False
cac.Visible = False
cname.Visible = False
CNedit.Visible = False
CFName.Visible = False
CFedit.Visible = False
CMName.Visible = False
CMedit.Visible = False
CDOB.Visible = False
CDedit.Visible = False
CGender.Visible = False
CGedit.Visible = False
CMobile.Visible = False
CMoedit.Visible = False
CMail.Visible = False
CMaedit.Visible = False
CLocation.Visible = False
CLoedit.Visible = False
Caadh.Visible = False
Caadedit.Visible = False

Enm.Visible = False
Enm1.Visible = False
Enm2.Visible = False
Enm3.Visible = False
Enm4.Visible = False
Enm5.Visible = False
Enm6.Visible = False
Enm7.Visible = False
TBD.Visible = False
TBD1.Visible = False
TBD2.Visible = False
TBD3.Visible = False

SEmpBtn.Visible = False
EmpNm.Visible = False
EmpPosition.Visible = False
EmpDOJ.Visible = False
EmpEID.Visible = False
EmpPhone.Visible = False
EmpMail.Visible = False
EmpSal.Visible = False
EmpAdd.Visible = False
StatementList.Visible = False


Make_Connection
con.Open cn
zc.Open " select * from customer_data order by acno", cn, adOpenDynamic, adLockOptimistic

If zc.EOF <> False Then
CID.Text = BCode.Caption + "11000000001"
ACNo = "10011111111"
Else
zc.MoveLast
kt = zc!Id
ju = Trim(Right(kt, 11)) + 1
CID.Text = BCode.Caption + ju
acn = Trim(zc!ACNo) + 1
ACNo.Text = acn
End If
zc.Close
con.Close


End Sub

Private Sub OpenProfile_Click()
Profile.Visible = True
Cross.Visible = True
LName.Visible = True
Pid.Visible = True
Pdoj.Visible = True
Pphone.Visible = True
Pmail.Visible = True
Paddress.Visible = True
PImg.Visible = True

SearchTran.Visible = False
BankBal.Visible = False
Stdep.Visible = False
Stwid.Visible = False
Sttra.Visible = False
Sbrnm.Visible = False
Scod.Visible = False
Sifs.Visible = False
Sintr.Visible = False
Sloc.Visible = False
TransactionList.Visible = False

SCustom.Visible = False
Bal.Visible = False
wamt.Visible = False
Wbtn.Visible = False
StatementList.Visible = False

Saving.Visible = False
Current.Visible = False
FD.Visible = False
CID.Visible = False
OBal.Visible = False
OName.Visible = False
FName.Visible = False
MName.Visible = False
DOB.Visible = False
DTPicker1.Visible = False
Gender.Visible = False
Male.Visible = False
Female.Visible = False
Transgender.Visible = False
Mobile.Visible = False
Email.Visible = False
Address.Visible = False
PLogout.Visible = False
LogYes.Visible = False
LogNo.Visible = False
HomeCover.Visible = False
APImg.Visible = False
UploadPImg.Visible = False
ASignImg.Visible = False
ASignUpload.Visible = False
APimgtext.Visible = False
ASignText.Visible = False
SearchBox.Visible = False
SearchCustomer.Visible = False
CPrev.Visible = False
CNext.Visible = False
CUpdate.Visible = False
CDel.Visible = False
CPic.Visible = False
CSign.Visible = False
CPicedit.Visible = False
CSignedit.Visible = False
ccid.Visible = False
ctype.Visible = False
cifsc.Visible = False
cac.Visible = False
cname.Visible = False
CNedit.Visible = False
CFName.Visible = False
CFedit.Visible = False
CMName.Visible = False
CMedit.Visible = False
CDOB.Visible = False
CDedit.Visible = False
CGender.Visible = False
CGedit.Visible = False
CMobile.Visible = False
CMoedit.Visible = False
CMail.Visible = False
CMaedit.Visible = False
CLocation.Visible = False
CLoedit.Visible = False
Caadh.Visible = False
Caadedit.Visible = False
ACancel.Visible = False
LOpen.Visible = False
FDPer.Visible = False

Enm.Visible = False
Enm1.Visible = False
Enm2.Visible = False
Enm3.Visible = False
Enm4.Visible = False
Enm5.Visible = False
Enm6.Visible = False
Enm7.Visible = False

SEmpBtn.Visible = False
EmpNm.Visible = False
EmpPosition.Visible = False
EmpDOJ.Visible = False
EmpEID.Visible = False
EmpPhone.Visible = False
EmpMail.Visible = False
EmpSal.Visible = False
EmpAdd.Visible = False

SearchS.Visible = False
SObal.Visible = False
STDp.Visible = False
STWd.Visible = False
SCbal.Visible = False
SAcNo.Visible = False
SAtyp.Visible = False
SAIfsc.Visible = False
SIra.Visible = False
SNom.Visible = False

End Sub

Private Sub Saving_Click()
FDPer.Visible = False
End Sub

Private Sub SCustom_Click()
Dim sfren As New ADODB.Recordset
Dim kks As String
Dim sts1 As String

Make_Connection
con.Open cn

sfren.Open " select * from customer_data order by acno", cn, adOpenDynamic, adLockOptimistic
kks = LCase(Trim(SearchBox.Text))
    
    f = 0
    Do While sfren.EOF <> True
    
    sts1 = LCase(Trim(sfren!Id))
    If kks = sts1 Then
    f = 1
ccid.Text = sfren!ACNo
ctype.Text = sfren!Name
cifsc.Text = sfren!FName
cac.Text = sfren!Mobile
Bal.Caption = sfren!Balance
CPic.Picture = LoadPicture(sfren!PImg)
CSign.Picture = LoadPicture(sfren!sImg)
Exit Do
End If
sfren.MoveNext
Loop
If f = 0 Then
MsgBox "No Record", vbExclamation
End If
con.Close
End Sub

Private Sub SearchCustomer_Click()
Dim sre As New ADODB.Recordset
Dim kks As String
Dim sts1 As String
cname.Enabled = False
CFName.Enabled = False
CMName.Enabled = False
CDOB.Enabled = False
CGender.Enabled = False
CMobile.Enabled = False
CMail.Enabled = False
CLocation.Enabled = False
Caadh.Enabled = False

Make_Connection
con.Open cn

sre.Open " select * from customer_data order by acno", cn, adOpenDynamic, adLockOptimistic
kks = LCase(Trim(SearchBox.Text))
     
    f = 0
    Do While sre.EOF <> True
   
    sts1 = LCase(Trim(sre!Id))
    If kks = sts1 Then
    f = 1
ccid.Text = sre!Id
ctype.Text = sre!actype
cifsc.Text = sre!IFSC
cac.Text = sre!ACNo
cname.Text = sre!Name
CFName.Text = sre!FName
CMName.Text = sre!MName
CDOB.Text = sre!DOB
CGender.Text = sre!Gender
CMobile.Text = sre!Mobile
CMail.Text = sre!Email
CLocation.Text = sre!Address
CPic.Picture = LoadPicture(sre!PImg)
CPictext.Text = sre!PImg
CSign.Picture = LoadPicture(sre!sImg)
CSigntext.Text = sre!sImg
If IsNull(sre!aadhar) Then
Caadh.Text = ""
MsgBox "KYC Pending, Update Aadhar Data", vbInformation
Caadh.Enabled = True
Caadh.SetFocus
Else
Caadh.Text = sre!aadhar
End If
Exit Do
End If
sre.MoveNext
Loop
If f = 0 Then
MsgBox "No Record", vbExclamation
End If
con.Close
End Sub

Private Sub SearchS_Click()
' Search Customer Statement

Dim stre As New ADODB.Recordset
Dim cstre As New ADODB.Recordset
Dim skks As String
Dim stas1 As String
Dim lista As New ADODB.Recordset
Dim tstat As Integer
Dim stai As Integer
Dim sttype As String
Dim stas2 As String
StatementList.ColumnHeaders.Clear
StatementList.ListItems.Clear

Make_Connection
con.Open cn

stre.Open " select * from customer_data order by acno", cn, adOpenDynamic, adLockOptimistic

skks = LCase(Trim(SearchBox.Text))
     
    f = 0
    Do While stre.EOF <> True
   
    stas1 = LCase(Trim(stre!Id))
    If skks = stas1 Then
    f = 1
    
SObal.Caption = stre!obalance
STDp.Caption = stre!tdeposit
STWd.Caption = stre!twithdrawl
SCbal.Caption = stre!Balance
SAcNo.Caption = stre!ACNo
SAtyp.Caption = stre!actype
SAIfsc.Caption = stre!IFSC
SIra.Caption = stre!interestrate

stas2 = Trim(UCase(SearchBox.Text))
lista.CursorLocation = adUseClient

lista.Open " select * from customer_statement where id = '" & stas2 & "' order by tid", cn, adOpenDynamic, adLockOptimistic

tstat = lista.RecordCount




    Dim itmx As ListItem ' Create a variable to add ListItem objects.
    Dim clmX As ColumnHeader ' Create an object variable for the ColumnHeader object.
   ' Add ColumnHeaders.
    Set clmX = StatementList.ColumnHeaders.Add(, , "Date & Time", StatementList.Width / 4, lvwColumnLeft)
    Set clmX = StatementList.ColumnHeaders.Add(, , "Transaction Detail", StatementList.Width / 3.25, lvwColumnCenter)
    Set clmX = StatementList.ColumnHeaders.Add(, , "Amount", StatementList.Width / 5, lvwColumnCenter)
    Set clmX = StatementList.ColumnHeaders.Add(, , "Available Balance", StatementList.Width / 4, lvwColumnCenter)
    
    StatementList.BorderStyle = ccFixedSingle ' Set BorderStyle property.
    StatementList.View = lvwReport ' Set View property to Report.
    
    For stai = 1 To tstat
    If lista!debit = "0" Then
sttype = "Deposit"
Else
If lista!credit = "0" Then
sttype = "Withdraw"
End If
End If
    ' Add a main item
    Set itmx = StatementList.ListItems.Add(, , Format(lista!dates, "dd-mm-yyyy HH:MM:SS"))
    ' Add two subitems for that item
    itmx.SubItems(1) = lista!tId & " (" & lista!tmode & " )" & "'" & sttype & "'"
    itmx.SubItems(2) = "Rs. " & lista!credit + lista!debit
    itmx.SubItems(3) = "Rs. " & lista!Balance
    
    lista.MoveNext
    Next stai

Exit Do
End If
stre.MoveNext
Loop
If f = 0 Then
MsgBox "No Record", vbExclamation
End If
con.Close
End Sub

Private Sub SearchTran_Click()
Dim stbt As New ADODB.Recordset
Dim i As Integer
Dim strdo As Integer
Make_Connection
con.Open cn
stbt.CursorLocation = adUseClient
stbt.Open " select tid from bank_transaction", cn, adOpenDynamic, adLockOptimistic
i = stbt.RecordCount
strdo = 1
Do While strdo <= i
If TransactionList.ListItems.Item(strdo).SubItems(1) = SearchBox.Text Then
TransactionList.ListItems.Item(strdo).Selected = True
TransactionList.SetFocus
Else
TransactionList.ListItems.Item(strdo).Selected = False
End If
strdo = strdo + 1
Loop
con.Close
End Sub

Private Sub SEmpBtn_Click()
Dim sempb As New ADODB.Recordset
Dim kemp As String
Dim semp As String

Make_Connection
con.Open cn

sempb.Open " select * from admin_login order by id", cn, adOpenDynamic, adLockOptimistic
kemp = LCase(Trim(SearchBox.Text))
    
    f = 0
    Do While sempb.EOF <> True
    semp = LCase(Trim(sempb!Id))
    If kemp = semp Then
    f = 1
    
EmpNm.Text = sempb!Name
EmpPosition.Text = sempb!post
EmpDOJ.Text = sempb!doj
EmpEID.Text = sempb!Id
EmpPhone.Text = sempb!phone
EmpMail.Text = sempb!Email
EmpSal.Text = sempb!salary
EmpAdd.Text = sempb!Address

Exit Do
End If
sempb.MoveNext
Loop
If f = 0 Then
MsgBox "No Record", vbExclamation
End If
con.Close
EmpNm.Locked = True
EmpPosition.Locked = True
EmpDOJ.Locked = True
EmpPhone.Locked = True
EmpMail.Locked = True
EmpSal.Locked = True
EmpAdd.Locked = True

EmpNm.FontBold = False
EmpPosition.FontBold = False
EmpDOJ.FontBold = False
EmpPhone.FontBold = False
EmpMail.FontBold = False
EmpSal.FontBold = False
EmpAdd.FontBold = False

End Sub


Private Sub StatementList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Call SortListView(StatementList, ColumnHeader.Index)
End Sub

Private Sub StatementPage_Click()
Set Me.Picture = Me.StatementImg
SearchBox.ToolTipText = "Enter ID"

SObal.Caption = ""
STDp.Caption = ""
STWd.Caption = ""
SCbal.Caption = ""
SAcNo.Caption = ""
SAtyp.Caption = ""
SAIfsc.Caption = ""
SIra.Caption = ""
SNom.Caption = ""
StatementList.ColumnHeaders.Clear
StatementList.ListItems.Clear

SObal.Visible = True
STDp.Visible = True
STWd.Visible = True
SCbal.Visible = True
SAcNo.Visible = True
SAtyp.Visible = True
SAIfsc.Visible = True
SIra.Visible = True
SNom.Visible = True
StatementList.Visible = True

SearchTran.Visible = False
BankBal.Visible = False
Stdep.Visible = False
Stwid.Visible = False
Sttra.Visible = False
Sbrnm.Visible = False
Scod.Visible = False
Sifs.Visible = False
Sintr.Visible = False
Sloc.Visible = False
TransactionList.Visible = False

Profile.Visible = False
Cross.Visible = False
LName.Visible = False
Pid.Visible = False
Pdoj.Visible = False
Pphone.Visible = False
Pmail.Visible = False
Paddress.Visible = False
PImg.Visible = False
ACancel.Visible = False
LOpen.Visible = False
FDPer.Visible = False

SCustom.Visible = False
Bal.Visible = False
wamt.Visible = False
Wbtn.Visible = False

WithdrawBTN.Visible = False
DepositBtn.Visible = False

Saving.Visible = False
Current.Visible = False
FD.Visible = False
CID.Visible = False
OBal.Visible = False
OName.Visible = False
FName.Visible = False
MName.Visible = False
DOB.Visible = False
DTPicker1.Visible = False
Gender.Visible = False
Male.Visible = False
Female.Visible = False
Transgender.Visible = False
Mobile.Visible = False
Email.Visible = False
Address.Visible = False
PLogout.Visible = False
LogYes.Visible = False
LogNo.Visible = False
Hifsc.Visible = False
HomeCover.Visible = False
BCode.Visible = False
Hmbl.Visible = False
Hmail.Visible = False
Haddress.Visible = False
Twitter.Visible = False
Linkedin.Visible = False
facebook.Visible = False
Instagram.Visible = False
Website.Visible = False
Download1.Visible = False
Download2.Visible = False
APImg.Visible = False
UploadPImg.Visible = False
ASignImg.Visible = False
ASignUpload.Visible = False
APimgtext.Visible = False
ASignText.Visible = False

SearchBox.Visible = True
SearchS.Visible = True
SearchCustomer.Visible = False
CPrev.Visible = False
CNext.Visible = False
CUpdate.Visible = False
CDel.Visible = False
CPic.Visible = False
CSign.Visible = False
CPicedit.Visible = False
CSignedit.Visible = False
ccid.Visible = False
ctype.Visible = False
cifsc.Visible = False
cac.Visible = False
cname.Visible = False
CNedit.Visible = False
CFName.Visible = False
CFedit.Visible = False
CMName.Visible = False
CMedit.Visible = False
CDOB.Visible = False
CDedit.Visible = False
CGender.Visible = False
CGedit.Visible = False
CMobile.Visible = False
CMoedit.Visible = False
CMail.Visible = False
CMaedit.Visible = False
CLocation.Visible = False
CLoedit.Visible = False
Caadh.Visible = False
Caadedit.Visible = False

Enm.Visible = False
Enm1.Visible = False
Enm2.Visible = False
Enm3.Visible = False
Enm4.Visible = False
Enm5.Visible = False
Enm6.Visible = False
Enm7.Visible = False
TBD.Visible = False
TBD1.Visible = False
TBD2.Visible = False
TBD3.Visible = False

SEmpBtn.Visible = False
EmpNm.Visible = False
EmpPosition.Visible = False
EmpDOJ.Visible = False
EmpEID.Visible = False
EmpPhone.Visible = False
EmpMail.Visible = False
EmpSal.Visible = False
EmpAdd.Visible = False

End Sub

Private Sub TransactionList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Call SortListView(TransactionList, ColumnHeader.Index)
End Sub

Private Sub TransactionPage_Click()
Set Me.Picture = Me.TransactionImg
SearchBox.ToolTipText = "Enter TID"
Dim todyt As New ADODB.Recordset
Dim libtra As New ADODB.Recordset
Dim ttodyt As Integer
Dim fedat As String
Dim stai1 As Long
Dim tottr As Long
Dim tottr1 As Long
tottr1 = 0
Dim tottr2 As Long
Dim tottr3 As Long
tottr3 = 0
Dim tbtran As Integer
Dim trlai As Long
TransactionList.ColumnHeaders.Clear
TransactionList.ListItems.Clear

SearchTran.Visible = True
BankBal.Visible = True
Stdep.Visible = True
Stwid.Visible = True
Sttra.Visible = True
Sbrnm.Visible = True
Scod.Visible = True
Sifs.Visible = True
Sintr.Visible = True
Sloc.Visible = True
TransactionList.Visible = True

Profile.Visible = False
Cross.Visible = False
LName.Visible = False
Pid.Visible = False
Pdoj.Visible = False
Pphone.Visible = False
Pmail.Visible = False
Paddress.Visible = False
PImg.Visible = False
ACancel.Visible = False
LOpen.Visible = False
FDPer.Visible = False

SCustom.Visible = False
Bal.Visible = False
wamt.Visible = False
Wbtn.Visible = False
StatementList.Visible = False

Saving.Visible = False
Current.Visible = False
FD.Visible = False
CID.Visible = False
OBal.Visible = False
OName.Visible = False
FName.Visible = False
MName.Visible = False
DOB.Visible = False
DTPicker1.Visible = False
Gender.Visible = False
Male.Visible = False
Female.Visible = False
Transgender.Visible = False
Mobile.Visible = False
Email.Visible = False
Address.Visible = False
PLogout.Visible = False
LogYes.Visible = False
LogNo.Visible = False
Hifsc.Visible = False
HomeCover.Visible = False
BCode.Visible = False
Hmbl.Visible = False
Hmail.Visible = False
Haddress.Visible = False
Twitter.Visible = False
Linkedin.Visible = False
facebook.Visible = False
Instagram.Visible = False
Website.Visible = False
Download1.Visible = False
Download2.Visible = False
APImg.Visible = False
UploadPImg.Visible = False
ASignImg.Visible = False
ASignUpload.Visible = False
APimgtext.Visible = False
ASignText.Visible = False

WithdrawBTN.Visible = False
DepositBtn.Visible = False

SearchBox.Visible = True
SearchCustomer.Visible = False
CPrev.Visible = False
CNext.Visible = False
CUpdate.Visible = False
CDel.Visible = False
CPic.Visible = False
CSign.Visible = False
CPicedit.Visible = False
CSignedit.Visible = False
ccid.Visible = False
ctype.Visible = False
cifsc.Visible = False
cac.Visible = False
cname.Visible = False
CNedit.Visible = False
CFName.Visible = False
CFedit.Visible = False
CMName.Visible = False
CMedit.Visible = False
CDOB.Visible = False
CDedit.Visible = False
CGender.Visible = False
CGedit.Visible = False
CMobile.Visible = False
CMoedit.Visible = False
CMail.Visible = False
CMaedit.Visible = False
CLocation.Visible = False
CLoedit.Visible = False
Caadh.Visible = False
Caadedit.Visible = False

Enm.Visible = False
Enm1.Visible = False
Enm2.Visible = False
Enm3.Visible = False
Enm4.Visible = False
Enm5.Visible = False
Enm6.Visible = False
Enm7.Visible = False
TBD.Visible = False
TBD1.Visible = False
TBD2.Visible = False
TBD3.Visible = False

SEmpBtn.Visible = False
EmpNm.Visible = False
EmpPosition.Visible = False
EmpDOJ.Visible = False
EmpEID.Visible = False
EmpPhone.Visible = False
EmpMail.Visible = False
EmpSal.Visible = False
EmpAdd.Visible = False

SearchS.Visible = False
SObal.Visible = False
STDp.Visible = False
STWd.Visible = False
SCbal.Visible = False
SAcNo.Visible = False
SAtyp.Visible = False
SAIfsc.Visible = False
SIra.Visible = False
SNom.Visible = False
con.Open
'Add Total Bank Balance Code Start
banbal.Open " select * from admin_login", cn, adOpenDynamic, adLockOptimistic
If IsNull(banbal!bank_bal) Then
con.Execute ("update admin_login set bank_bal= 0")
con.Execute ("commit")
Else
BankBal.Caption = banbal!bank_bal
End If
banbal.Close
'Total Deposit code
banbal.Open
If IsNull(banbal!tdeposit) Then
con.Execute ("update admin_login set tdeposit= 0")
con.Execute ("commit")
banbal.Close
banbal.Open " select * from admin_login", cn, adOpenDynamic, adLockOptimistic
Stdep.Caption = banbal!tdeposit
Else
Stdep.Caption = banbal!tdeposit
End If
banbal.Close

'Total Withdrawl Code
banbal.Open
If IsNull(banbal!twithdrawl) Then
con.Execute ("update admin_login set twithdrawl= 0")
con.Execute ("commit")
banbal.Close
banbal.Open " select * from admin_login", cn, adOpenDynamic, adLockOptimistic
Stwid.Caption = banbal!twithdrawl
Else
Stwid.Caption = banbal!twithdrawl
End If
banbal.Close
'Code End

'Today Transaction Code Start
fedat = Format(Now, "dd-mmm-yyyy")

todyt.CursorLocation = adUseClient
todyt.Open " select * from customer_statement where trunc(Dates) = '" & fedat & "' order by tid", cn, adOpenDynamic, adLockOptimistic

ttodyt = todyt.RecordCount
For stai1 = 1 To ttodyt
    If IsNull(todyt!ttype) Or ttodyt = "0" Then
Sttra.Caption = "No Transaction"
Else
If todyt!ttype = "Credit" Then
tottr = todyt!credit
tottr1 = tottr1 + tottr
Else
If todyt!ttype = "Debit" Then
tottr2 = todyt!debit
tottr3 = tottr3 + tottr2
End If
End If

End If
todyt.MoveNext
    Next stai1
    Sttra.Caption = "C = " & tottr1 & " | D = " & tottr3
'Code End
Sbrnm.Caption = Label5.Caption
Scod.Caption = BCode.Caption
Sifs.Caption = Hifsc.Caption
Sintr.Caption = "S: 4% | C: 2.5% | FD: 6%"
Sloc.Caption = Paddress.Caption

'Transaction List Code
libtra.CursorLocation = adUseClient

libtra.Open " select * from Bank_Transaction order by tid", cn, adOpenDynamic, adLockOptimistic

tbtran = libtra.RecordCount




    Dim itmx As ListItem ' Create a variable to add ListItem objects.
    Dim clmX As ColumnHeader ' Create an object variable for the ColumnHeader object.
   ' Add ColumnHeaders.
    Set clmX = TransactionList.ColumnHeaders.Add(, , "Date & Time", TransactionList.Width / 7.5, lvwColumnLeft)
    Set clmX = TransactionList.ColumnHeaders.Add(, , "Transaction ID", TransactionList.Width / 6, lvwColumnCenter)
    Set clmX = TransactionList.ColumnHeaders.Add(, , "Transaction Type", TransactionList.Width / 4.5, lvwColumnCenter)
    Set clmX = TransactionList.ColumnHeaders.Add(, , "ID", TransactionList.Width / 6, lvwColumnCenter)
    Set clmX = TransactionList.ColumnHeaders.Add(, , "Name", TransactionList.Width / 6, lvwColumnCenter)
    Set clmX = TransactionList.ColumnHeaders.Add(, , "Amount", TransactionList.Width / 7.5, lvwColumnCenter)
    Set clmX = TransactionList.ColumnHeaders.Add(, , "A/C No.", TransactionList.Width / 6, lvwColumnCenter)
    Set clmX = TransactionList.ColumnHeaders.Add(, , "IFSC", TransactionList.Width / 6, lvwColumnCenter)
    
    TransactionList.BorderStyle = ccFixedSingle ' Set BorderStyle property.
    TransactionList.View = lvwReport ' Set View property to Report.
    
    For trlai = 1 To tbtran
    ' Add a main item
    Set itmx = TransactionList.ListItems.Add(, , Format(libtra!dates, "dd-mm-yyyy HH:MM:SS"))
    ' Add subitems for that item
    itmx.SubItems(1) = libtra!tId
    itmx.SubItems(2) = libtra!ttype & " | " & libtra!tmode
    itmx.SubItems(3) = libtra!Id
    itmx.SubItems(4) = libtra!Name
    itmx.SubItems(5) = libtra!amount
    itmx.SubItems(6) = libtra!ACNo
    itmx.SubItems(7) = libtra!IFSC
    
    libtra.MoveNext
    Next trlai

con.Close
End Sub

Private Sub UploadPImg_Click()
CD1.Filter = "JPG File | *.jpg|GIF File|*.gif"
CD1.ShowOpen
If CD1.FileName <> "" Then
APImg.Picture = LoadPicture(CD1.FileName)
APimgtext.Text = CD1.FileName
End If

End Sub
Private Sub wamt_KeyPress(KeyAscii As Integer)
If KeyAscii <> 48 And KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 52 And KeyAscii <> 53 And KeyAscii <> 54 And KeyAscii <> 55 And KeyAscii <> 56 And KeyAscii <> 57 And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 46 Then
    KeyAscii = 0
    MsgBox ("Please Enter Numbers Only")
    End If
End Sub

Private Sub WithdrawBTN_Click()
Dim sfres As New ADODB.Recordset
Dim wduid As String
Dim dwduid As String

Dim wdifsc As String
Dim wdacno As String
Dim wdid As String
Dim wdname As String
Dim wdmobile As String
Dim wddate As String
Dim wddebit As String
Dim wdcredit As String
Dim wdbalance As String
Dim vgbj As String
Dim twid As String
Dim wtmode As String
Dim vtdep2 As String
Dim wtypfi As String
vgbj = ccid.Text
wtmode = "Cash"
wtypfi = "Debit"

If ccid.Text = "" Then
MsgBox ("Please Search Customer")
SearchBox.SetFocus
Else
If wamt.Text = "" Then
MsgBox ("Please Enter Amount")
wamt.SetFocus
Else
If Bal.Caption = 0 Then
MsgBox ("Low Balance"), vbCritical
Else
Make_Connection
con.Open cn
sfres.Open " select * from customer_data where acno='" & vgbj & "'", cn, adOpenDynamic, adLockOptimistic

wdifsc = sfres!IFSC
wdacno = sfres!ACNo
wdid = sfres!Id
wdname = sfres!Name
wdmobile = sfres!Mobile
wddate = Now()
wddebit = wamt.Text
wdcredit = "0"
wdbalance = Trim(sfres!Balance - wamt.Text)
wdtid = sfres!ltid + 1
twid = Trim(sfres!twithdrawl + wamt.Text)

con.Execute (" ALTER SESSION SET NLS_DATE_FORMAT='dd-mm-yyyy hh:mi:ss AM'")
con.Execute ("insert into customer_statement (ifsc,acno,id,name,mobile,dates,debit,credit,balance,tid,tmode,ttype)values('" & wdifsc & "','" & wdacno & "','" & wdid & "','" & wdname & "','" & wdmobile & "','" & wddate & "','" & wddebit & "','" & wdcredit & "','" & wdbalance & "','" & wdtid & "','" & wtmode & "','" & wtypfi & "')")
con.Execute ("insert into Bank_Transaction (dates,tid,ttype,tmode,id,name,amount,acno,ifsc)values('" & wddate & "','" & wdtid & "','" & wtypfi & "','" & wtmode & "','" & wdid & "','" & wdname & "','" & wddebit & "','" & wdacno & "','" & wdifsc & "')")
con.Execute ("commit")

con.Execute ("update customer_data set ltid='" & wdtid & "',balance='" & wdbalance & "',twithdrawl='" & twid & "' where id='" & wdid & "' ")
con.Execute ("commit")

'Add Total Bank Balance Code Start
banbal.Open " select * from admin_login", cn, adOpenDynamic, adLockOptimistic
If IsNull(banbal!bank_bal) Then
con.Execute ("update admin_login set bank_bal= 0")
con.Execute ("commit")
banbal.Close
banbal.Open " select * from admin_login", cn, adOpenDynamic, adLockOptimistic
bbala = Trim(banbal!bank_bal - wamt.Text)
Else
bbala = Trim(banbal!bank_bal - wamt.Text)
End If
con.Execute ("update admin_login set bank_bal='" & bbala & "'")
con.Execute ("commit")
banbal.Close

banbal.Open
If IsNull(banbal!twithdrawl) Then
con.Execute ("update admin_login set twithdrawl= 0")
con.Execute ("commit")
banbal.Close
banbal.Open " select * from admin_login", cn, adOpenDynamic, adLockOptimistic
vtdep2 = Trim(banbal!twithdrawl + wamt.Text)
con.Execute ("update admin_login set twithdrawl= '" & vtdep2 & "'")
con.Execute ("commit")
Else
vtdep2 = Trim(banbal!twithdrawl + wamt.Text)
con.Execute ("update admin_login set twithdrawl= '" & vtdep2 & "'")
con.Execute ("commit")
End If
banbal.Close
'Code End
MsgBox ("Withdrawn Successfully, Your Balance is " + wdbalance)

Bal.Caption = wdbalance
wamt.Text = ""

con.Close
End If
End If
End If
End Sub
Private Sub WithdrawPage_Click()
Set Me.Picture = Me.WithdrawImg
SearchBox.ToolTipText = "Enter ID"
SCustom.Visible = True
Bal.Visible = True
wamt.Visible = True
Wbtn.Visible = True
WithdrawBTN.Visible = True
DepositBtn.Visible = False
StatementList.Visible = False

ccid.Text = ""
ctype.Text = ""
cifsc.Text = ""
cac.Text = ""
Bal.Caption = ""

CPic.Picture = LoadPicture()
CSign.Picture = LoadPicture()

Profile.Visible = False
Cross.Visible = False
LName.Visible = False
Pid.Visible = False
Pdoj.Visible = False
Pphone.Visible = False
Pmail.Visible = False
Paddress.Visible = False
PImg.Visible = False
ACancel.Visible = False
LOpen.Visible = False
FDPer.Visible = False

SearchTran.Visible = False
BankBal.Visible = False
Stdep.Visible = False
Stwid.Visible = False
Sttra.Visible = False
Sbrnm.Visible = False
Scod.Visible = False
Sifs.Visible = False
Sintr.Visible = False
Sloc.Visible = False
TransactionList.Visible = False

Saving.Visible = False
Current.Visible = False
FD.Visible = False
CID.Visible = False
OBal.Visible = False
OName.Visible = False
FName.Visible = False
MName.Visible = False
DOB.Visible = False
DTPicker1.Visible = False
Gender.Visible = False
Male.Visible = False
Female.Visible = False
Transgender.Visible = False
Mobile.Visible = False
Email.Visible = False
Address.Visible = False
PLogout.Visible = False
LogYes.Visible = False
LogNo.Visible = False
Hifsc.Visible = False
HomeCover.Visible = False
BCode.Visible = False
Hmbl.Visible = False
Hmail.Visible = False
Haddress.Visible = False
Twitter.Visible = False
Linkedin.Visible = False
facebook.Visible = False
Instagram.Visible = False
Website.Visible = False
Download1.Visible = False
Download2.Visible = False
APImg.Visible = False
UploadPImg.Visible = False
ASignImg.Visible = False
ASignUpload.Visible = False
APimgtext.Visible = False
ASignText.Visible = False

SearchS.Visible = False
SObal.Visible = False
STDp.Visible = False
STWd.Visible = False
SCbal.Visible = False
SAcNo.Visible = False
SAtyp.Visible = False
SAIfsc.Visible = False
SIra.Visible = False
SNom.Visible = False

SearchBox.Visible = True
SearchCustomer.Visible = False
CPrev.Visible = False
CNext.Visible = False
CUpdate.Visible = False
CDel.Visible = False
CPic.Visible = True
CSign.Visible = True
CPicedit.Visible = False
CSignedit.Visible = False
ccid.Visible = True
ctype.Visible = True
cifsc.Visible = True
cac.Visible = True
cname.Visible = False
CNedit.Visible = False
CFName.Visible = False
CFedit.Visible = False
CMName.Visible = False
CMedit.Visible = False
CDOB.Visible = False
CDedit.Visible = False
CGender.Visible = False
CGedit.Visible = False
CMobile.Visible = False
CMoedit.Visible = False
CMail.Visible = False
CMaedit.Visible = False
CLocation.Visible = False
CLoedit.Visible = False
Caadh.Visible = False
Caadedit.Visible = False

Enm.Visible = False
Enm1.Visible = False
Enm2.Visible = False
Enm3.Visible = False
Enm4.Visible = False
Enm5.Visible = False
Enm6.Visible = False
Enm7.Visible = False
TBD.Visible = False
TBD1.Visible = False
TBD2.Visible = False
TBD3.Visible = False

SEmpBtn.Visible = False
EmpNm.Visible = False
EmpPosition.Visible = False
EmpDOJ.Visible = False
EmpEID.Visible = False
EmpPhone.Visible = False
EmpMail.Visible = False
EmpSal.Visible = False
EmpAdd.Visible = False

End Sub
Private Sub SearchBox_KeyPress(KeyAscii As Integer)
If SCustom.Visible = True Then

If KeyAscii = 13 Then
SCustom_Click
End If
Else
If SearchCustomer.Visible = True Then
If KeyAscii = 13 Then
SearchCustomer_Click
End If
Else
If SEmpBtn.Visible = True Then
If KeyAscii = 13 Then
SEmpBtn_Click
End If
Else
If SearchTran.Visible = True Then
If KeyAscii = 13 Then
SearchTran_Click
End If
Else
If SearchS.Visible = True Then
If KeyAscii = 13 Then
SearchS_Click

End If
End If
End If
End If
End If
End If
End Sub

