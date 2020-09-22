VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4680
   ClientLeft      =   2595
   ClientTop       =   2400
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   2200
      Left            =   15
      Top             =   4200
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "By: Asim Shafiq Baig"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00764E10&
      Height          =   465
      Left            =   480
      TabIndex        =   1
      Top             =   4380
      Width           =   5805
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Boutique Management System"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00764E10&
      Height          =   645
      Left            =   435
      TabIndex        =   0
      Top             =   3375
      Width           =   5685
   End
   Begin VB.Image Image1 
      Height          =   4110
      Left            =   420
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   225
      Width           =   5910
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
    login.Show
    Unload Me
End Sub
