VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form customer 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BMS - Customers"
   ClientHeight    =   5640
   ClientLeft      =   1485
   ClientTop       =   2100
   ClientWidth     =   8820
   FillColor       =   &H00764E10&
   BeginProperty Font 
      Name            =   "Monotype Corsiva"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00764E10&
   Icon            =   "customer.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   8820
   Begin TabDlg.SSTab SSTab1 
      Height          =   3945
      Left            =   45
      TabIndex        =   4
      Top             =   1350
      Width           =   8700
      _ExtentX        =   15346
      _ExtentY        =   6959
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   16777215
      ForeColor       =   9992024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Create New Customer"
      TabPicture(0)   =   "customer.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Shape1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Line5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label6"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label8"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label9"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtaccount"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdlogin"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtname"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtphones"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Command1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Settle Customer Dues"
      TabPicture(1)   =   "customer.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Shape2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label5"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Line6"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label10"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label11"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Command2"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txtdues"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Command3"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cmbname"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cmbaccount"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Command4"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).ControlCount=   11
      Begin VB.CommandButton Command4 
         BackColor       =   &H00E2D1D3&
         Caption         =   "<< Get Dues"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -68280
         MaskColor       =   &H00987758&
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2070
         UseMaskColor    =   -1  'True
         Width           =   1575
      End
      Begin VB.ComboBox cmbaccount 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00987758&
         Height          =   360
         Left            =   -70815
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1575
         Width           =   1395
      End
      Begin VB.ComboBox cmbname 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00987758&
         Height          =   360
         Left            =   -69390
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1575
         Width           =   2715
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E2D1D3&
         Caption         =   "Save Dues"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -70822
         MaskColor       =   &H00987758&
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2745
         UseMaskColor    =   -1  'True
         Width           =   1635
      End
      Begin VB.TextBox txtdues 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00987758&
         Height          =   360
         Left            =   -70822
         MaxLength       =   99
         TabIndex        =   18
         Top             =   2115
         Width           =   2385
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E2D1D3&
         Caption         =   "&cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   -69142
         MaskColor       =   &H00987758&
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2730
         UseMaskColor    =   -1  'True
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E2D1D3&
         Caption         =   "&cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5700
         MaskColor       =   &H00987758&
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2910
         UseMaskColor    =   -1  'True
         Width           =   1575
      End
      Begin VB.TextBox txtphones 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00987758&
         Height          =   360
         Left            =   4020
         TabIndex        =   9
         Top             =   2340
         Width           =   3285
      End
      Begin VB.TextBox txtname 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00987758&
         Height          =   360
         Left            =   4020
         MaxLength       =   99
         TabIndex        =   8
         Top             =   1785
         Width           =   3285
      End
      Begin VB.CommandButton cmdlogin 
         BackColor       =   &H00E2D1D3&
         Caption         =   "&Create Account"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4020
         MaskColor       =   &H00987758&
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2925
         UseMaskColor    =   -1  'True
         Width           =   1635
      End
      Begin VB.TextBox txtaccount 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00764E10&
         Height          =   360
         Left            =   4020
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1230
         Width           =   1575
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Select Customer:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00987758&
         Height          =   285
         Left            =   -73620
         TabIndex        =   23
         Top             =   1635
         Width           =   2745
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Dues: (Rs.)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00987758&
         Height          =   285
         Left            =   -73635
         TabIndex        =   22
         Top             =   2220
         Width           =   2745
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Phone Nos.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00987758&
         Height          =   285
         Left            =   1200
         TabIndex        =   15
         Top             =   2400
         Width           =   2745
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00987758&
         Height          =   285
         Left            =   1200
         TabIndex        =   14
         Top             =   1860
         Width           =   2745
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Account No.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00987758&
         Height          =   285
         Left            =   1215
         TabIndex        =   12
         Top             =   1275
         Width           =   2745
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00764E10&
         BorderWidth     =   2
         X1              =   -74895
         X2              =   -66435
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Settle Customer Dues"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00764E10&
         Height          =   315
         Left            =   -74865
         TabIndex        =   6
         Top             =   420
         Width           =   8445
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00764E10&
         BorderWidth     =   2
         X1              =   120
         X2              =   8580
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Create New Customer Account"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00764E10&
         Height          =   315
         Left            =   150
         TabIndex        =   5
         Top             =   420
         Width           =   8445
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00764E10&
         Height          =   3465
         Left            =   -74925
         Top             =   390
         Width           =   8535
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00764E10&
         Height          =   3465
         Left            =   45
         Top             =   390
         Width           =   8535
      End
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Account No.:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00987758&
      Height          =   285
      Left            =   1680
      TabIndex        =   13
      Top             =   2805
      Width           =   2745
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00764E10&
      BorderWidth     =   2
      X1              =   5895
      X2              =   6315
      Y1              =   1245
      Y2              =   1065
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00764E10&
      BorderWidth     =   2
      X1              =   6285
      X2              =   8790
      Y1              =   1065
      Y2              =   1065
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00764E10&
      BorderWidth     =   2
      X1              =   60
      X2              =   8790
      Y1              =   1275
      Y2              =   1275
   End
   Begin VB.Label loginbar 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   6210
      TabIndex        =   3
      Top             =   1050
      Width           =   2565
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Beta Version 1.0"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00987758&
      Height          =   225
      Left            =   6930
      TabIndex        =   2
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "CopyRights 2003 - All Rights Reserved to Asim Shafiq Baig (Developer BMS)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00987758&
      Height          =   195
      Left            =   90
      TabIndex        =   1
      Top             =   5415
      Width           =   8655
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00764E10&
      BorderWidth     =   2
      X1              =   30
      X2              =   8760
      Y1              =   5370
      Y2              =   5370
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Boutique Management System"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00764E10&
      Height          =   675
      Left            =   2520
      TabIndex        =   0
      Top             =   405
      Width           =   7245
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   1470
      Left            =   15
      Picture         =   "customer.frx":047A
      Stretch         =   -1  'True
      Top             =   30
      Width           =   2520
   End
End
Attribute VB_Name = "customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbaccount_Click()
    cmbname.ListIndex = cmbaccount.ListIndex
    txtdues.Text = ""
End Sub

Private Sub cmbname_Click()
    cmbaccount.ListIndex = cmbname.ListIndex
    txtdues.Text = ""
End Sub

Private Sub cmdlogin_Click()
    If txtname.Text = "" Then
        MsgBox "Enter A Name Please:", vbCritical
        txtname.SetFocus
        Exit Sub
    End If
        
    If Not validity2(txtname, "Name") Then
        txtname.SetFocus
        Exit Sub
    End If
    
    If Not validity2(txtphones, "Phones") Then
        txtphones.SetFocus
        Exit Sub
    End If
    
    Call openconn
    sqlstr = "insert into customer values('" & txtaccount.Text & "','" & txtname.Text & "','" & txtphones.Text & "',0)"
    Call rs(sqlstr)
    Call closeconn
    
    MsgBox "Customer Account Created Successfully!", vbInformation
        
    Unload Me
    
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    If (cmbaccount.ListIndex = -1) Then
        MsgBox "Please Select A Customer!", vbCritical
        cmbaccount.SetFocus
    Exit Sub
    End If
    
    If txtdues.Text = "" Then
        MsgBox "Please Enter Dues or Enter 0 in Dues", vbCritical
        txtdues.SetFocus
        Exit Sub
    End If
        
    
    If Not IsNumeric(txtdues.Text) Then
        MsgBox "Customer Dues Have to be Number Format!", vbCritical
        txtdues.SetFocus
    Exit Sub
    End If
    
    If CCur(txtdues.Text) < 0 Then
        MsgBox "Cannot Enter Negative Values In Dues!", vbCritical
        txtdues.SetFocus
        Exit Sub
    End If
    
        
    Call openconn
    
    sqlstr = "update customer set dues=" & CCur(txtdues.Text) & " where customerid = '" & cmbaccount.Text & "' and name = '" & cmbname.Text & "'"
    
    Call rs(sqlstr)
    
    Call closeconn
    
    MsgBox "Customer Dues Saved Successfully!", vbInformation
        
    Unload Me
    
    
End Sub

Private Sub Command4_Click()
    If (cmbaccount.ListIndex = -1) Then
        MsgBox "Please Select A Customer!", vbCritical
        cmbaccount.SetFocus
    Exit Sub
    End If
    
    Call openconn
    sqlstr = "select dues from customer where customerid='" & cmbaccount.Text & "' and name = '" & cmbname.Text & "'"
    
    Call rs(sqlstr)
    
    If adoRS.EOF Then
        MsgBox "No Customer Selected", vbCritical
    Exit Sub
    End If
        
    txtdues.Text = adoRS.Fields(0)
    Call closeconn
        
    
End Sub

Private Sub Form_Load()
    loginbar.Caption = "User = " & userid
    
    Call openconn
    sqlstr = "select customerid from customer"
    Call rs(sqlstr)
   
        
    If (adoRS.EOF) Then
        txtaccount = "000001"
    Else
        hello = adoRS.Sort
        adoRS.MoveLast
        txtaccount = Right("000000" & CStr(CLng(adoRS.Fields(0)) + 1), 6)
    End If
    
    Call closeconn
    
    
    Call openconn
    sqlstr = "select customerid, name from customer"
    Call rs(sqlstr)
    
    If Not (adoRS.EOF) Then
        hello = adoRS.Sort
        adoRS.MoveFirst
        While Not (adoRS.EOF)
            cmbaccount.AddItem adoRS.Fields("customerid")
            cmbname.AddItem adoRS.Fields("name")
            adoRS.MoveNext
        Wend
    End If
    
    Call closeconn
    
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MainForm2.Show
End Sub

