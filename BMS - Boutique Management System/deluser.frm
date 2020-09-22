VERSION 5.00
Begin VB.Form deluser 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BMS - Create New User"
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
   Icon            =   "deluser.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   8820
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Delete System Login Users"
      ForeColor       =   &H00764E10&
      Height          =   2625
      Left            =   2745
      TabIndex        =   7
      Top             =   1950
      Width           =   4905
      Begin VB.TextBox txtadmin 
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
         Top             =   1320
         Width           =   2745
      End
      Begin VB.TextBox txtlogin 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00987758&
         Height          =   315
         Left            =   1785
         MaxLength       =   15
         TabIndex        =   1
         ToolTipText     =   "Enter LogIn ID Here"
         Top             =   675
         Width           =   2745
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E2D1D3&
         Caption         =   "Cancel"
         Height          =   375
         Left            =   3165
         MaskColor       =   &H00987758&
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1950
         UseMaskColor    =   -1  'True
         Width           =   1335
      End
      Begin VB.CommandButton cmdlogin 
         BackColor       =   &H00E2D1D3&
         Caption         =   "&Delete"
         Default         =   -1  'True
         Height          =   375
         Left            =   1800
         MaskColor       =   &H00987758&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1950
         UseMaskColor    =   -1  'True
         Width           =   1335
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Admin Password:"
         ForeColor       =   &H00987758&
         Height          =   285
         Left            =   0
         TabIndex        =   9
         Top             =   1350
         Width           =   1755
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Login Id:"
         ForeColor       =   &H00987758&
         Height          =   285
         Left            =   345
         TabIndex        =   8
         Top             =   705
         Width           =   1425
      End
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
      Picture         =   "deluser.frx":0442
      Stretch         =   -1  'True
      Top             =   30
      Width           =   2520
   End
End
Attribute VB_Name = "deluser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

    
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
        
       
    If txtadmin.Text = "" Then
        MsgBox "Enter An Administrator Password:", vbCritical
        txtadmin.SetFocus
        Exit Sub
    End If
    
    If Not validity(txtadmin, "Administrator Password") Then
        txtadmin.SetFocus
        Exit Sub
    End If
    
    Call openconn
    sqlstr = "select * from users where loginid = '" & Crypt("administrator") & "' and pass = '" & Crypt(txtadmin.Text) & "'"
    Call rs(sqlstr)
            
    If (adoRS.EOF) Then
        MsgBox "Wrong Administrator Password! Try Again With Correct Administrator Password", vbCritical
        txtadmin.Text = ""
        txtadmin.SetFocus
        Call closeconn
        Exit Sub
    End If
    
    Call closeconn
    
    Call openconn
    sqlstr = "select * from users where loginid = '" & Crypt(txtlogin.Text) & "'"
    Call rs(sqlstr)
            
    If (adoRS.EOF) Then
        MsgBox "Login ID (" & txtlogin.Text & ") Does Not Exist! Enter Valid Login ID", vbCritical
        txtlogin.Text = ""
        txtlogin.SetFocus
        Call closeconn
        Exit Sub
    End If
    
    Call closeconn
    
   
    Call openconn
    sqlstr = "delete from users where loginid = '" & Crypt(txtlogin.Text) & "'"
    
    Call rs(sqlstr)
    Call closeconn
    
    MsgBox "User Deleted Successfully!", vbInformation
    
    Unload Me
End Sub


Private Sub Command1_Click()
    Unload Me
End Sub

