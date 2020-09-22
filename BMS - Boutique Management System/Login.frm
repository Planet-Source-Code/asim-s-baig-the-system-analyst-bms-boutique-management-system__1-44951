VERSION 5.00
Begin VB.Form login 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BMS - System Login Screen"
   ClientHeight    =   5640
   ClientLeft      =   1485
   ClientTop       =   2100
   ClientWidth     =   8820
   FillColor       =   &H00764E10&
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00764E10&
   Icon            =   "Login.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   8820
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "System Login"
      ForeColor       =   &H00764E10&
      Height          =   2445
      Left            =   2730
      TabIndex        =   7
      Top             =   2025
      Width           =   4905
      Begin VB.TextBox txtlogin 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00987758&
         Height          =   315
         Left            =   1785
         MaxLength       =   15
         TabIndex        =   1
         ToolTipText     =   "Enter LogIn ID Here"
         Top             =   495
         Width           =   2745
      End
      Begin VB.TextBox txtpass 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00987758&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1785
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   2
         ToolTipText     =   "Enter Password Here"
         Top             =   1095
         Width           =   2745
      End
      Begin VB.CommandButton cmdlogin 
         BackColor       =   &H00E2D1D3&
         Caption         =   "&Login"
         Default         =   -1  'True
         Height          =   375
         Left            =   1800
         MaskColor       =   &H00987758&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1710
         UseMaskColor    =   -1  'True
         Width           =   1335
      End
      Begin VB.CommandButton cmdcancel 
         BackColor       =   &H00E2D1D3&
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   3180
         MaskColor       =   &H00987758&
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1710
         UseMaskColor    =   -1  'True
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Login Id:"
         ForeColor       =   &H00987758&
         Height          =   285
         Left            =   300
         TabIndex        =   12
         Top             =   540
         Width           =   1425
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         ForeColor       =   &H00987758&
         Height          =   285
         Left            =   270
         TabIndex        =   11
         Top             =   1140
         Width           =   1425
      End
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Delete User"
      ForeColor       =   &H00764E10&
      Height          =   255
      Left            =   4185
      MouseIcon       =   "Login.frx":0442
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   4785
      Width           =   1785
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "|"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00764E10&
      Height          =   375
      Left            =   4920
      TabIndex        =   10
      Top             =   4485
      Width           =   165
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Change User Password"
      ForeColor       =   &H00764E10&
      Height          =   255
      Left            =   5025
      MouseIcon       =   "Login.frx":074C
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   4500
      Width           =   2325
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Create New User"
      ForeColor       =   &H00764E10&
      Height          =   255
      Left            =   3120
      MouseIcon       =   "Login.frx":0A56
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   4500
      Width           =   1785
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
      TabIndex        =   6
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
      TabIndex        =   5
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
      Top             =   375
      Width           =   7245
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   3735
      Left            =   1680
      Shape           =   2  'Oval
      Top             =   1425
      Width           =   7035
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00987758&
      FillStyle       =   0  'Solid
      Height          =   4065
      Left            =   45
      Shape           =   2  'Oval
      Top             =   1260
      Width           =   8730
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   1470
      Left            =   15
      Picture         =   "Login.frx":0D60
      Stretch         =   -1  'True
      Top             =   30
      Width           =   2520
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdlogin_Click()
    If txtlogin.Text = "" Then
        MsgBox "Enter A Login ID:", vbCritical
        txtlogin.SetFocus
        Exit Sub
    End If
        
    If Not validity(txtlogin, "Login ID") Then
        txtlogin.SetFocus
        Exit Sub
    End If
        
    If txtpass.Text = "" Then
        MsgBox "Enter A Password:", vbCritical
        txtpass.SetFocus
        Exit Sub
    End If
    
    If Not validity(txtpass, "Password") Then
        txtpass.SetFocus
        Exit Sub
    End If
    
    Call openconn
    sqlstr = "select * from users where loginid = '" & Crypt(txtlogin.Text) & "' and pass = '" & Crypt(txtpass.Text) & "'"
    Call rs(sqlstr)
            
    If (adoRS.EOF) Then
        MsgBox "Wrong Login ID or Password! Try Again", vbCritical
        txtlogin.Text = ""
        txtpass.Text = ""
        txtlogin.SetFocus
        Call closeconn
        Exit Sub
    End If
    
    Call closeconn
    
    userid = txtlogin.Text
    
    Unload Me
    MainForm.Show
    
End Sub

Private Sub Label6_Click()
    newuser.Show vbModal
End Sub

Private Sub Label7_Click()
    changepass.Show vbModal
End Sub

Private Sub Label9_Click()
    deluser.Show vbModal
End Sub

Private Sub Form_Unload(Cancel As Integer)
    endform.Show
End Sub

