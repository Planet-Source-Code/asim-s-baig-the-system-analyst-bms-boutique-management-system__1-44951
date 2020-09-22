VERSION 5.00
Begin VB.Form endform 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3795
   ClientLeft      =   2745
   ClientTop       =   2580
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(Email): helloworld@yahoo.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00764E10&
      Height          =   270
      Left            =   1305
      TabIndex        =   6
      Top             =   3225
      Width           =   4125
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(Ph): (9221) 1234567 - 1234567"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00764E10&
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   2940
      Width           =   4125
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(Address): B-10, Phillip Avenue, jkhdfjkdkfghkjfhdkj gkjfdh, Karachi, Pakistan."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00764E10&
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   2430
      Width           =   4125
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "S   y   s   t   e   m       A   n   a   l   y   s   t"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00987758&
      Height          =   285
      Left            =   555
      TabIndex        =   3
      Top             =   2115
      Width           =   5655
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Contact"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   480
      TabIndex        =   2
      Top             =   1440
      Width           =   5805
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A s i m   S h a f i q   B a i g"
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
      Height          =   465
      Left            =   510
      TabIndex        =   1
      Top             =   1845
      Width           =   5805
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Boutique Management System"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00764E10&
      Height          =   765
      Left            =   1410
      TabIndex        =   0
      Top             =   495
      Width           =   5955
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00764E10&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   375
      Left            =   540
      Top             =   1410
      Width           =   5655
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00764E10&
      FillColor       =   &H00E0E0E0&
      Height          =   2175
      Left            =   540
      Top             =   1395
      Width           =   5655
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   30
      Picture         =   "endform.frx":0000
      Stretch         =   -1  'True
      Top             =   15
      Width           =   2220
   End
End
Attribute VB_Name = "endform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
