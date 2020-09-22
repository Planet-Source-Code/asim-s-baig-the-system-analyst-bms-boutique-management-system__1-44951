VERSION 5.00
Begin VB.Form changepass 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BMS - Change Password Screen"
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
   Icon            =   "changepass.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   8820
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Change System User's Password:"
      ForeColor       =   &H00764E10&
      Height          =   2625
      Left            =   2760
      TabIndex        =   10
      Top             =   1935
      Width           =   4905
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Admin?"
         ForeColor       =   &H00987758&
         Height          =   285
         Left            =   3570
         TabIndex        =   3
         Top             =   885
         Width           =   1125
      End
      Begin VB.CommandButton cmdlogin 
         BackColor       =   &H00E2D1D3&
         Caption         =   "&Change"
         Default         =   -1  'True
         Height          =   375
         Left            =   1800
         MaskColor       =   &H00987758&
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2145
         UseMaskColor    =   -1  'True
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E2D1D3&
         Caption         =   "Cancel"
         Height          =   375
         Left            =   3165
         MaskColor       =   &H00987758&
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2145
         UseMaskColor    =   -1  'True
         Width           =   1335
      End
      Begin VB.TextBox txtpass 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00987758&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2160
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   2
         ToolTipText     =   "Enter Password Here"
         Top             =   840
         Width           =   1275
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
         Top             =   405
         Width           =   2745
      End
      Begin VB.TextBox txtpass2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00987758&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1785
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   4
         ToolTipText     =   "Enter Password Here"
         Top             =   1275
         Width           =   2745
      End
      Begin VB.TextBox txtpass3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00987758&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1785
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   5
         ToolTipText     =   "Enter Password Here"
         Top             =   1725
         Width           =   2745
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Current/Admin Pass:"
         ForeColor       =   &H00987758&
         Height          =   285
         Left            =   -30
         TabIndex        =   14
         Top             =   870
         Width           =   2130
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Login Id:"
         ForeColor       =   &H00987758&
         Height          =   285
         Left            =   345
         TabIndex        =   13
         Top             =   435
         Width           =   1425
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "New Pass:"
         ForeColor       =   &H00987758&
         Height          =   285
         Left            =   60
         TabIndex        =   12
         Top             =   1305
         Width           =   1665
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Pass:"
         ForeColor       =   &H00987758&
         Height          =   285
         Left            =   0
         TabIndex        =   11
         Top             =   1755
         Width           =   1755
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
      TabIndex        =   9
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
      TabIndex        =   8
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
      Left            =   1710
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
      Picture         =   "changepass.frx":0442
      Stretch         =   -1  'True
      Top             =   30
      Width           =   2520
   End
End
Attribute VB_Name = "changepass"
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
        
    If txtpass.Text = "" Then
        MsgBox "Enter current or Administrator's Password:", vbCritical
        txtpass.SetFocus
        Exit Sub
    End If
    
    If Not validity(txtpass, "Current / Administrator's Password") Then
        txtpass.SetFocus
        Exit Sub
    End If
    
    If txtpass2.Text = "" Then
        MsgBox "Enter A New Password:", vbCritical
        txtpass2.SetFocus
        Exit Sub
    End If
    
    If Not validity(txtpass2, "New Password") Then
        txtpass2.SetFocus
        Exit Sub
    End If
    
    
    If txtpass3.Text = "" Then
        MsgBox "Enter A Confirmation Password:", vbCritical
        txtpass3.SetFocus
        Exit Sub
    End If
    
    If Not validity(txtpass3, "Confirmation Password") Then
        txtpass3.SetFocus
        Exit Sub
    End If
    
    If txtpass2.Text <> txtpass3.Text Then
        MsgBox "Password & Confirm Passwords Donot Match!", vbCritical
        txtpass2.Text = ""
        txtpass3.Text = ""
        txtpass2.SetFocus
        Exit Sub
    End If
    
    If Check1.Value = 0 Then
        
        Call openconn
        sqlstr = "select * from users where loginid = '" & Crypt(txtlogin.Text) & "'"
        Call rs(sqlstr)
                
        If (adoRS.EOF) Then
            MsgBox "Login ID Does Not Exist! Enter Correct Login ID", vbCritical
            txtlogin.Text = ""
            txtlogin.SetFocus
            Call closeconn
            Exit Sub
        End If
        
        Call closeconn
    End If
    
    Call openconn
    
    If Check1.Value = 1 Then
        sqlstr = "select * from users where loginid = '" & Crypt("administrator") & "' and pass = '" & Crypt(txtpass.Text) & "'"
    Else
        sqlstr = "select * from users where loginid = '" & Crypt(txtlogin.Text) & "' and pass = '" & Crypt(txtpass.Text) & "'"
    End If
    
    Call rs(sqlstr)
            
    If (adoRS.EOF) Then
        If Check1.Value = 1 Then
            MsgBox "Wrong Administrator Password! Try Again With Correct Administrator Password", vbCritical
        Else
            MsgBox "Wrong Current Password! Try Again With Correct Current Password", vbCritical
        End If
        
        txtpass.Text = ""
        txtpass.SetFocus
        Call closeconn
        Exit Sub
    End If
    
    Call closeconn
        
    Call openconn
    sqlstr = "update users set pass = '" & Crypt(txtpass2.Text) & "' where loginid = '" & Crypt(txtlogin.Text) & "'"
    
    Call rs(sqlstr)
    
    Call closeconn
    
    MsgBox "Password Changed Successfully!", vbInformation
    
    Unload Me

End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

