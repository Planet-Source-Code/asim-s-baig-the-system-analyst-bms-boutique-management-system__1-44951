VERSION 5.00
Begin VB.Form MainForm2 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BMS - SYSTEM CONTROL PANEL"
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
   Icon            =   "MainForm2.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   8820
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Exit System"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00764E10&
      Height          =   375
      Left            =   7065
      MouseIcon       =   "MainForm2.frx":0442
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   4980
      Width           =   1575
   End
   Begin VB.Label lblreports 
      BackStyle       =   0  'Transparent
      Caption         =   "View Reports"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00764E10&
      Height          =   345
      Left            =   3825
      MouseIcon       =   "MainForm2.frx":074C
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   4065
      Width           =   2265
   End
   Begin VB.Label lblcustomers 
      BackStyle       =   0  'Transparent
      Caption         =   "Customers"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00764E10&
      Height          =   345
      Left            =   3840
      MouseIcon       =   "MainForm2.frx":0A56
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   3420
      Width           =   1875
   End
   Begin VB.Label lblsale 
      BackStyle       =   0  'Transparent
      Caption         =   "Clothes Selling"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00764E10&
      Height          =   345
      Left            =   3855
      MouseIcon       =   "MainForm2.frx":0D60
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   2760
      Width           =   2715
   End
   Begin VB.Label lblpurchase 
      BackStyle       =   0  'Transparent
      Caption         =   "New/Modify Clothes"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00764E10&
      Height          =   345
      Left            =   3885
      MouseIcon       =   "MainForm2.frx":106A
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2055
      Width           =   3495
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
   Begin VB.Image Image5 
      Height          =   390
      Left            =   2955
      MouseIcon       =   "MainForm2.frx":1374
      MousePointer    =   99  'Custom
      Picture         =   "MainForm2.frx":167E
      Stretch         =   -1  'True
      Top             =   4065
      Width           =   795
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   2955
      MouseIcon       =   "MainForm2.frx":1BFB
      MousePointer    =   99  'Custom
      Picture         =   "MainForm2.frx":1F05
      Stretch         =   -1  'True
      Top             =   3405
      Width           =   795
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   2970
      MouseIcon       =   "MainForm2.frx":2482
      MousePointer    =   99  'Custom
      Picture         =   "MainForm2.frx":278C
      Stretch         =   -1  'True
      Top             =   2745
      Width           =   780
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   2985
      MouseIcon       =   "MainForm2.frx":2D09
      MousePointer    =   99  'Custom
      Picture         =   "MainForm2.frx":3013
      Stretch         =   -1  'True
      Top             =   2055
      Width           =   780
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
      Picture         =   "MainForm2.frx":3590
      Stretch         =   -1  'True
      Top             =   30
      Width           =   2520
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00764E10&
      BorderWidth     =   2
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   405
      Left            =   7065
      Shape           =   4  'Rounded Rectangle
      Top             =   4920
      Width           =   1695
   End
End
Attribute VB_Name = "MainForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
    frmmain2.Show
End Sub


Private Sub Form_Unload(Cancel As Integer)
    endform.Show
End Sub

Private Sub Image2_Click()
    Call lblpurchase_Click
End Sub

Private Sub Image3_Click()
    Call lblsale_Click
End Sub

Private Sub Image4_Click()
    Call lblcustomers_Click
End Sub

Private Sub Image5_Click()
    Call lblreports_Click
End Sub

Private Sub Label4_Click()
    Unload Me
End Sub

Private Sub lblcustomers_Click()
    Unload Me
    customer.Show
End Sub

Private Sub lblpurchase_Click()
    Unload Me
    purchase.Show
End Sub

Private Sub lblreports_Click()
    Unload Me
    reports.Show
End Sub

Private Sub lblsale_Click()
    Unload Me
    sale.Show
End Sub
