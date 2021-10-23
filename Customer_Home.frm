VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Customer_Home 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dashboard"
   ClientHeight    =   10365
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   19110
   Icon            =   "Customer_Home.frx":0000
   LinkTopic       =   "Customer_Home"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Customer_Home.frx":10CA
   ScaleHeight     =   691
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1274
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   480
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
            Picture         =   "Customer_Home.frx":A2BA
            Key             =   "UpArrow"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Customer_Home.frx":B394
            Key             =   "DownArrow"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView TransactionList 
      Height          =   4455
      Left            =   810
      TabIndex        =   13
      Top             =   5520
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   7858
      View            =   3
      LabelEdit       =   1
      SortOrder       =   -1  'True
      Sorted          =   -1  'True
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
   Begin VB.TextBox TAmt 
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
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   12480
      TabIndex        =   6
      Top             =   8580
      Width           =   3255
   End
   Begin VB.TextBox TAcNo 
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
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   12480
      TabIndex        =   5
      Top             =   7845
      Width           =   3120
   End
   Begin VB.TextBox TIFSC 
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
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   12480
      TabIndex        =   4
      Top             =   7125
      Width           =   3255
   End
   Begin VB.Label RefreshTran 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   9120
      MouseIcon       =   "Customer_Home.frx":C46E
      MousePointer    =   99  'Custom
      TabIndex        =   20
      ToolTipText     =   "Refresh"
      Top             =   5040
      Width           =   495
   End
   Begin VB.Label TReset 
      BackStyle       =   0  'Transparent
      Height          =   360
      Left            =   18300
      MouseIcon       =   "Customer_Home.frx":D538
      MousePointer    =   99  'Custom
      TabIndex        =   19
      ToolTipText     =   "Reset"
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label TSNm 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Rubik"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   300
      Left            =   15000
      TabIndex        =   18
      Top             =   6570
      Width           =   3015
      WordWrap        =   -1  'True
   End
   Begin VB.Label TSAcn 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   15600
      MouseIcon       =   "Customer_Home.frx":E602
      MousePointer    =   99  'Custom
      TabIndex        =   17
      ToolTipText     =   "Search"
      Top             =   7800
      Width           =   495
   End
   Begin VB.Label DepositBtn 
      Height          =   375
      Left            =   16560
      TabIndex        =   16
      Top             =   9720
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label TMoney 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   13800
      MouseIcon       =   "Customer_Home.frx":F6CC
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   9480
      Width           =   2175
   End
   Begin VB.Label Logout 
      BackStyle       =   0  'Transparent
      Height          =   525
      Left            =   18225
      MouseIcon       =   "Customer_Home.frx":10796
      MousePointer    =   99  'Custom
      TabIndex        =   14
      ToolTipText     =   "Logout"
      Top             =   1170
      Width           =   630
   End
   Begin VB.Image CPic 
      Height          =   1770
      Left            =   16545
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   1770
   End
   Begin VB.Label Mail 
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
      Height          =   300
      Left            =   12360
      TabIndex        =   12
      Top             =   5220
      Width           =   3735
   End
   Begin VB.Label Mobile 
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
      Height          =   300
      Left            =   12360
      TabIndex        =   11
      Top             =   4755
      Width           =   3735
   End
   Begin VB.Label IFSC 
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
      Height          =   300
      Left            =   12360
      TabIndex        =   10
      Top             =   4275
      Width           =   3735
   End
   Begin VB.Label Branch 
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
      Height          =   300
      Left            =   12360
      TabIndex        =   9
      Top             =   3795
      Width           =   3735
   End
   Begin VB.Label CNM 
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
      Height          =   300
      Left            =   12360
      TabIndex        =   8
      Top             =   3315
      Width           =   3735
   End
   Begin VB.Label CID 
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
      Height          =   300
      Left            =   12360
      TabIndex        =   7
      Top             =   2820
      Width           =   3735
   End
   Begin VB.Label ACNo 
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
      Left            =   1080
      TabIndex        =   3
      ToolTipText     =   "Account Number"
      Top             =   3120
      Width           =   3255
   End
   Begin VB.Label ACTyp 
      BackStyle       =   0  'Transparent
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
      TabIndex        =   2
      Top             =   2760
      Width           =   3975
   End
   Begin VB.Label Bal 
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
      Height          =   375
      Left            =   8160
      TabIndex        =   1
      Top             =   2820
      Width           =   1815
   End
   Begin VB.Label Unm 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   15360
      TabIndex        =   0
      Top             =   330
      Width           =   3615
   End
End
Attribute VB_Name = "Customer_Home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim banbal As New ADODB.Recordset
Dim st1 As String
Dim st2 As String
Dim tbtran As Integer
Dim sttype As String
Dim libtra As New ADODB.Recordset
Dim cn As String

Public Function Make_Connection()
 
    cn = "Provider=MSDAORA.1;User ID=bank/SABank;Data Source=localhost;Persist Security Info=False"
   
End Function
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
Private Sub Form_Load()
Dim kk As String

Make_Connection
    
If Login.Combo1.Text = "Customer" Then

rs.Open " select * from customer_data", cn, adOpenDynamic, adLockOptimistic
kk = LCase(Trim(Login.Id.Text))
    rs.MoveFirst
    f = 0
    Do While rs.EOF <> True
    st1 = LCase(Trim(rs!Id))
    If kk = st1 Then
    f = 1
Unm.Caption = "WELCOME " + UCase(rs!Name)
Bal.Caption = rs!Balance
ACTyp.Caption = "SABank " & rs!actype & " Account"
ACNo.Caption = rs!ACNo
CID.Caption = rs!Id
CNM.Caption = rs!Name
Branch.Caption = rs!Branch_name
IFSC.Caption = rs!IFSC
Mobile.Caption = rs!Mobile
Mail.Caption = rs!Email
CPic.Picture = LoadPicture(rs!PImg)
st2 = UCase(st1)

RefreshTran_Click
    
    Exit Do
End If
rs.MoveNext
Loop
If f = 0 Then
Unm.Caption = "No Record"
End If
rs.Close
End If
End Sub

Private Sub Logout_Click()
answer = MsgBox("Are you sure you want Logout ", vbCritical + vbYesNo, "Warning")
If answer = vbYes Then
Unload Me
Login.Show
End If
End Sub

Private Sub RefreshTran_Click()
TransactionList.ListItems.Clear
TransactionList.ColumnHeaders.Clear
'Transaction List Code
libtra.CursorLocation = adUseClient

libtra.Open " select * from customer_statement where id = '" & st2 & "' order by tid", cn, adOpenDynamic, adLockOptimistic

tbtran = libtra.RecordCount

    Dim itmx As ListItem ' Create a variable to add ListItem objects.
    Dim clmX As ColumnHeader ' Create an object variable for the ColumnHeader object.
   ' Add ColumnHeaders.
    Set clmX = TransactionList.ColumnHeaders.Add(, , "Transaction Detail", TransactionList.Width / 4, lvwColumnLeft)
    Set clmX = TransactionList.ColumnHeaders.Add(, , "Date & Time", TransactionList.Width / 3.25, lvwColumnCenter)
    Set clmX = TransactionList.ColumnHeaders.Add(, , "Amount", TransactionList.Width / 5, lvwColumnCenter)
    Set clmX = TransactionList.ColumnHeaders.Add(, , "Available Balance", TransactionList.Width / 4.8, lvwColumnCenter)
    
    TransactionList.BorderStyle = ccFixedSingle ' Set BorderStyle property.
    TransactionList.View = lvwReport ' Set View property to Report.
    
    For trlai = 1 To tbtran
    If libtra!debit = "0" Then
sttype = "Deposit"
Else
If libtra!credit = "0" Then
sttype = "Withdraw"
End If
End If
    ' Add a main item
    Set itmx = TransactionList.ListItems.Add(, , libtra!tId & " (" & libtra!tmode & " )" & "'" & sttype & "'")
    ' Add two subitems for that item
    itmx.SubItems(1) = Format(libtra!dates, "dd-mm-yyyy HH:MM:SS")
    itmx.SubItems(2) = "Rs. " & libtra!credit + libtra!debit
    itmx.SubItems(3) = "Rs. " & libtra!Balance
    
    
    
    libtra.MoveNext
    Next trlai
    libtra.Close
End Sub

Private Sub TAcNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TSAcn_Click
End If
If KeyAscii <> 48 And KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 52 And KeyAscii <> 53 And KeyAscii <> 54 And KeyAscii <> 55 And KeyAscii <> 56 And KeyAscii <> 57 And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 46 Then
    KeyAscii = 0
    MsgBox ("Please Enter Numbers Only")
    End If
End Sub

Private Sub TAmt_KeyPress(KeyAscii As Integer)
If KeyAscii <> 48 And KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 52 And KeyAscii <> 53 And KeyAscii <> 54 And KeyAscii <> 55 And KeyAscii <> 56 And KeyAscii <> 57 And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 46 Then
    KeyAscii = 0
    MsgBox ("Please Enter Numbers Only")
    End If
End Sub

Private Sub TReset_Click()
TSNm.Caption = ""
TAcNo.Text = ""
TIFSC.Text = ""
TIFSC.Enabled = True
TAmt.Text = ""
End Sub

Private Sub TSAcn_Click()
Dim mt1 As String
Dim mt2 As String
Dim mt3 As String
Dim mt4 As String
Dim mrs As New ADODB.Recordset

mrs.Open " select ifsc,acno,name from customer_data", cn, adOpenDynamic, adLockOptimistic
mt1 = Trim(TAcNo.Text)
mt3 = UCase(Trim(TIFSC.Text))
    mrs.MoveFirst
    f = 0
    Do While mrs.EOF <> True
    mt2 = Trim(mrs!ACNo)
    mt4 = UCase(Trim(mrs!IFSC))
    If mt1 = mt2 And mt3 = mt4 Then
    f = 1
    TSNm.Caption = mrs!Name
    TIFSC.Enabled = False
TAmt.SetFocus
Exit Do
End If
mrs.MoveNext
Loop
If f = 0 Then
TSNm.Caption = ""
MsgBox "IFSC & Account No. Not Found", vbInformation
End If
mrs.Close
End Sub

Private Sub TMoney_Click()
Dim sfres As New ADODB.Recordset
Dim chacn As New ADODB.Recordset
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
Dim smacn As String
Dim twid As String
Dim wtmode As String
Dim vtdep2 As String
Dim wtypfi As String
Dim chname As String

vgbj = ACNo.Caption
smacn = TAcNo.Text
wtmode = "Sent To " & TSNm.Caption & " A/C No. " & TAcNo.Text
wtypfi = "Debit"
If TSNm.Caption = "" Then
MsgBox "Search Customer First", vbCritical
Else
Make_Connection
con.Open cn
chacn.Open " select name from customer_data where acno='" & smacn & "'", cn, adOpenDynamic, adLockOptimistic
If chacn.EOF Then
MsgBox "Invalid Account Number", vbCritical
Else
chname = chacn!Name
End If
chacn.Close
con.Close
If Trim(UCase(chname)) = Trim(UCase(TSNm.Caption)) Then


If TSNm.Caption = "" Then
MsgBox ("Please Search Customer")
TIFSC.SetFocus
Else
If TAmt.Text = "" Then
MsgBox ("Please Enter Amount")
TAmt.SetFocus
Else
If Bal.Caption = 0 Then
MsgBox ("Low Balance"), vbCritical
Else
If Bal.Caption < TAmt.Text Then
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
wddebit = TAmt.Text
wdcredit = "0"

wdbalance = Trim(sfres!Balance - TAmt.Text)
wdtid = sfres!ltid + 1
twid = Trim(sfres!twithdrawl + TAmt.Text)

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
bbala = Trim(banbal!bank_bal - TAmt.Text)
Else
bbala = Trim(banbal!bank_bal - TAmt.Text)
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
vtdep2 = Trim(banbal!twithdrawl + TAmt.Text)
con.Execute ("update admin_login set twithdrawl= '" & vtdep2 & "'")
con.Execute ("commit")
Else
vtdep2 = Trim(banbal!twithdrawl + TAmt.Text)
con.Execute ("update admin_login set twithdrawl= '" & vtdep2 & "'")
con.Execute ("commit")
End If
banbal.Close
'Code End
MsgBox ("Your Balance is " + wdbalance), vbInformation, "Remaining Balance"

Bal.Caption = wdbalance

con.Close
DepositBTN_Click

End If
End If
End If
End If
Else
MsgBox "Search Customer Again", vbCritical
End If
End If
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

Vgbk = TAcNo.Text
tmode = "Received from " & CNM.Caption & " A/C No. " & ACNo.Caption

If TSNm.Caption = "" Then
MsgBox ("Please Search Customer")
TAcNo.SetFocus
Else
If TAmt.Text = "" Then
MsgBox ("Please Enter Amount")
TAmt.SetFocus
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
wdcredit = TAmt.Text
wdbalance = Trim(sfres!Balance + TAmt.Text)

If IsNull(banbal!bank_bal) Then
con.Execute ("update admin_login set bank_bal= 0")
con.Execute ("commit")
banbal.Close
banbal.Open " select bank_bal from admin_login", cn, adOpenDynamic, adLockOptimistic
bbala = Trim(banbal!bank_bal + TAmt.Text)
Else
bbala = Trim(banbal!bank_bal + TAmt.Text)
End If


wdtid = sfres!ltid + 1
tdep = Trim(sfres!tdeposit + TAmt.Text)
con.Execute (" ALTER SESSION SET NLS_DATE_FORMAT='dd-mm-yyyy hh:mi:ss AM'")
con.Execute ("insert into customer_statement (ifsc,acno,id,name,mobile,dates,credit,debit,balance,tid,tmode,ttype)values('" & wdifsc & "','" & wdacno & "','" & wdid & "','" & wdname & "','" & wdmobile & "','" & wddate & "','" & wdcredit & "','" & wddebit & "','" & wdbalance & "','" & wdtid & "','" & tmode & "','" & dtypfi & "')")
con.Execute ("insert into Bank_Transaction (dates,tid,ttype,tmode,id,name,amount,acno,ifsc)values('" & wddate & "','" & wdtid & "','" & dtypfi & "','" & tmode & "','" & wdid & "','" & wdname & "','" & wdcredit & "','" & wdacno & "','" & wdifsc & "')")
con.Execute ("commit")
con.Execute ("update admin_login set bank_bal='" & bbala & "'")
con.Execute ("update customer_data set ltid='" & wdtid & "',balance='" & wdbalance & "',tdeposit='" & tdep & "' where id='" & wdid & "' ")
con.Execute ("commit")

vtdep1 = Trim(banbal!tdeposit + TAmt.Text)
con.Execute ("update admin_login set tdeposit= '" & vtdep1 & "'")
con.Execute ("commit")
MsgBox ("Transferred Successfully, To " + wdname)

TAmt.Text = ""
TSNm.Caption = ""
TIFSC.Text = ""
TAcNo.Text = ""
TIFSC.Enabled = True
banbal.Close
con.Close
RefreshTran_Click
End If
End If
End Sub

Private Sub TransactionList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Call SortListView(TransactionList, ColumnHeader.Index)
End Sub
