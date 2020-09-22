VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Begin VB.Form MainForm 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BMS - Welcome"
   ClientHeight    =   5400
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
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   8820
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   465
      Left            =   7035
      TabIndex        =   2
      Top             =   4935
      Width           =   1830
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Skip Intro"
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
         Height          =   315
         Left            =   120
         MouseIcon       =   "MainForm.frx":0442
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   75
         Width           =   1575
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00764E10&
         BorderWidth     =   2
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   405
         Left            =   60
         Shape           =   4  'Rounded Rectangle
         Top             =   15
         Width           =   1695
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   30000
      Left            =   8265
      Top             =   1035
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   4155
      Left            =   45
      TabIndex        =   1
      Top             =   1230
      Width           =   8745
      _cx             =   15425
      _cy             =   7329
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   0   'False
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   0   'False
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
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
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   1470
      Left            =   15
      Picture         =   "MainForm.frx":074C
      Stretch         =   -1  'True
      Top             =   30
      Width           =   2520
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    ShockwaveFlash1.Movie = App.Path & "/BMS - Movie.swf"
    ShockwaveFlash1.LoadMovie 1, App.Path & "/BMS - Movie.swf"
    ShockwaveFlash1.Loop = False
    ShockwaveFlash1.Menu = False
End Sub

Private Sub Label4_Click()
    Unload Me
    MainForm2.Show
End Sub



Private Sub Timer1_Timer()
    MainForm2.Show
    Unload MainForm
End Sub


Private Sub Form_Unload(Cancel As Integer)
    endform.Show
End Sub

