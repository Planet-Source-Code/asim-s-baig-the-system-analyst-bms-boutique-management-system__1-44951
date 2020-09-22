VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form reports 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BMS - View Reports"
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
   Icon            =   "reports.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   8820
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   180
      Left            =   -1170
      TabIndex        =   0
      Top             =   1950
      Width           =   645
   End
   Begin VB.Frame frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2025
      Left            =   120
      TabIndex        =   12
      Top             =   3180
      Visible         =   0   'False
      Width           =   8550
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Only Customers WIth Dues"
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
         Height          =   405
         Left            =   4523
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   840
         Width           =   3045
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "All Customers"
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
         Height          =   405
         Left            =   1253
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   840
         Value           =   -1  'True
         Width           =   3045
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E2D1D3&
         Caption         =   "Show Statement"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3488
         MaskColor       =   &H00987758&
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1440
         UseMaskColor    =   -1  'True
         Width           =   1845
      End
      Begin MSComCtl2.DTPicker dtfrom 
         Height          =   360
         Left            =   165
         TabIndex        =   5
         Top             =   870
         Width           =   4170
         _ExtentX        =   7355
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarForeColor=   7753232
         CalendarTitleForeColor=   7753232
         Format          =   19595264
         CurrentDate     =   37654
      End
      Begin MSComCtl2.DTPicker dtto 
         Height          =   360
         Left            =   4515
         TabIndex        =   6
         Top             =   870
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarForeColor=   7753232
         CalendarTitleForeColor=   7753232
         Format          =   19595264
         CurrentDate     =   37654
      End
      Begin VB.Label lblmain 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Left            =   120
         TabIndex        =   15
         Top             =   210
         Width           =   8325
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Date To:"
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
         Left            =   4500
         TabIndex        =   14
         Top             =   600
         Width           =   1065
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Date From:"
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
         Left            =   180
         TabIndex        =   13
         Top             =   600
         Width           =   1065
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E2D1D3&
      Caption         =   ":: Customer Dues Statement ::"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2378
      MaskColor       =   &H00987758&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2625
      UseMaskColor    =   -1  'True
      Width           =   4065
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E2D1D3&
      Caption         =   ":: Profit And Losses Statement ::"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2378
      MaskColor       =   &H00987758&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2220
      UseMaskColor    =   -1  'True
      Width           =   4065
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E2D1D3&
      Caption         =   "::Selling And Income Statement ::"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2378
      MaskColor       =   &H00987758&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1830
      UseMaskColor    =   -1  'True
      Width           =   4065
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PRINT REPORTS"
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
      Left            =   225
      TabIndex        =   11
      Top             =   1380
      Width           =   8445
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00764E10&
      BorderWidth     =   2
      X1              =   180
      X2              =   8640
      Y1              =   1710
      Y2              =   1710
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00764E10&
      BorderWidth     =   2
      X1              =   5850
      X2              =   6270
      Y1              =   1245
      Y2              =   1065
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00764E10&
      BorderWidth     =   2
      X1              =   6240
      X2              =   8745
      Y1              =   1065
      Y2              =   1065
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00764E10&
      BorderWidth     =   2
      X1              =   15
      X2              =   8745
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
      Left            =   6165
      TabIndex        =   10
      Top             =   1050
      Width           =   2565
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
      TabIndex        =   9
      Top             =   375
      Width           =   7245
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
      TabIndex        =   8
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
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   1470
      Left            =   15
      Picture         =   "reports.frx":0442
      Stretch         =   -1  'True
      Top             =   30
      Width           =   2520
   End
End
Attribute VB_Name = "reports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    GL_REPORT = "PL"
    frame1.Visible = True
    lblmain.Caption = Command1.Caption
    Command4.FontBold = False
    Command1.FontBold = True
    Command2.FontBold = False
    Label4.Visible = True
    Label16.Visible = True
    dtfrom.Visible = True
    dtto.Visible = True
    Option1.Visible = False
    Option2.Visible = False
End Sub

Private Sub Command2_Click()
    GL_REPORT = "Customer"
    frame1.Visible = True
    lblmain.Caption = Command2.Caption
    Command4.FontBold = False
    Command1.FontBold = False
    Command2.FontBold = True
    Label4.Visible = False
    Label16.Visible = False
    dtfrom.Visible = False
    dtto.Visible = False
    Option1.Visible = True
    Option2.Visible = True
End Sub

Private Sub Command3_Click()
    frmprint.Show vbModal, Me
End Sub

Private Sub Command4_Click()
    GL_REPORT = "Selling"
    frame1.Visible = True
    lblmain.Caption = Command4.Caption
    Command4.FontBold = True
    Command1.FontBold = False
    Command2.FontBold = False
    Label4.Visible = True
    Label16.Visible = True
    dtfrom.Visible = True
    dtto.Visible = True
    Option1.Visible = False
    Option2.Visible = False
End Sub

Private Sub Form_Load()
   loginbar.Caption = "User = " & userid
   
   dtfrom.Value = Date
   dtto.Value = Date
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MainForm2.Show
End Sub

