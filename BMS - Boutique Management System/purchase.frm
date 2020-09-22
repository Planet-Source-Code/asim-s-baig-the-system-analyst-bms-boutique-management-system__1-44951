VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form purchase 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BMS - Clothes"
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
   Icon            =   "purchase.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   8820
   Begin TabDlg.SSTab SSTab1 
      Height          =   3945
      Left            =   45
      TabIndex        =   9
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
      TabCaption(0)   =   "New Clothes Entry"
      TabPicture(0)   =   "purchase.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Shape1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label9"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label8"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label6"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Line6"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label16"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Command1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtdescription"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtname"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdsave1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtserial"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtqty"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtcost"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Option1(1)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Option1(0)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "Modify / Change Clothes Entry"
      TabPicture(1)   =   "purchase.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmbdescription"
      Tab(1).Control(1)=   "cmbname"
      Tab(1).Control(2)=   "cmbserial"
      Tab(1).Control(3)=   "Command4"
      Tab(1).Control(4)=   "Frame1"
      Tab(1).Control(5)=   "Command3"
      Tab(1).Control(6)=   "txtdescription2"
      Tab(1).Control(7)=   "txtname2"
      Tab(1).Control(8)=   "cmdsave2"
      Tab(1).Control(9)=   "txtqty2"
      Tab(1).Control(10)=   "txtcost2"
      Tab(1).Control(11)=   "Label17"
      Tab(1).Control(12)=   "Line5"
      Tab(1).Control(13)=   "Label15"
      Tab(1).Control(14)=   "Label14"
      Tab(1).Control(15)=   "Label13"
      Tab(1).Control(16)=   "Label12"
      Tab(1).Control(17)=   "Label11"
      Tab(1).Control(18)=   "Label10"
      Tab(1).Control(19)=   "Shape2"
      Tab(1).ControlCount=   20
      Begin VB.ComboBox cmbdescription 
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
         Left            =   -69465
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   825
         Width           =   2985
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
         Left            =   -72210
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   825
         Width           =   2715
      End
      Begin VB.ComboBox cmbserial 
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
         Left            =   -73620
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   825
         Width           =   1395
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00E2D1D3&
         Caption         =   "<< Get Details"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -68100
         MaskColor       =   &H00987758&
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1260
         UseMaskColor    =   -1  'True
         Width           =   1575
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   435
         Left            =   -70800
         TabIndex        =   37
         Top             =   2970
         Width           =   3285
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "UnStiched"
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
            Index           =   1
            Left            =   1320
            TabIndex        =   30
            Top             =   90
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Stiched"
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
            Index           =   0
            Left            =   30
            TabIndex        =   29
            Top             =   90
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.CommandButton Command3 
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
         TabIndex        =   34
         Top             =   3390
         UseMaskColor    =   -1  'True
         Width           =   1575
      End
      Begin VB.TextBox txtdescription2 
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
         TabIndex        =   25
         Top             =   1695
         Width           =   3285
      End
      Begin VB.TextBox txtname2 
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
         MaxLength       =   49
         TabIndex        =   24
         Top             =   1260
         Width           =   2535
      End
      Begin VB.CommandButton cmdsave2 
         BackColor       =   &H00E2D1D3&
         Caption         =   "&Save Entry"
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
         TabIndex        =   32
         Top             =   3405
         UseMaskColor    =   -1  'True
         Width           =   1635
      End
      Begin VB.TextBox txtqty2 
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
         TabIndex        =   26
         Top             =   2145
         Width           =   1155
      End
      Begin VB.TextBox txtcost2 
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
         TabIndex        =   28
         Top             =   2595
         Width           =   1425
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Stiched"
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
         Index           =   0
         Left            =   4260
         TabIndex        =   5
         Top             =   3105
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "UnStiched"
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
         Index           =   1
         Left            =   5550
         TabIndex        =   6
         Top             =   3105
         Width           =   1215
      End
      Begin VB.TextBox txtcost 
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
         Left            =   4178
         TabIndex        =   4
         Top             =   2670
         Width           =   1425
      End
      Begin VB.TextBox txtqty 
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
         Left            =   4178
         TabIndex        =   3
         Top             =   2220
         Width           =   1155
      End
      Begin VB.TextBox txtserial 
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
         Left            =   4185
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   900
         Width           =   1575
      End
      Begin VB.CommandButton cmdsave1 
         BackColor       =   &H00E2D1D3&
         Caption         =   "&Save Entry"
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
         Left            =   4178
         MaskColor       =   &H00987758&
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3405
         UseMaskColor    =   -1  'True
         Width           =   1635
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
         Left            =   4178
         MaxLength       =   49
         TabIndex        =   1
         Top             =   1335
         Width           =   3285
      End
      Begin VB.TextBox txtdescription 
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
         Left            =   4178
         TabIndex        =   2
         Top             =   1770
         Width           =   3285
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
         Left            =   5858
         MaskColor       =   &H00987758&
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3390
         UseMaskColor    =   -1  'True
         Width           =   1575
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Modify / Change Clothes Entry"
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
         TabIndex        =   40
         Top             =   435
         Width           =   8445
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00764E10&
         BorderWidth     =   2
         X1              =   -74895
         X2              =   -66435
         Y1              =   765
         Y2              =   765
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "New Clothes Entry"
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
         Left            =   135
         TabIndex        =   39
         Top             =   405
         Width           =   8445
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00764E10&
         BorderWidth     =   2
         X1              =   105
         X2              =   8565
         Y1              =   735
         Y2              =   735
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Select Cloth:"
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
         Left            =   -74955
         TabIndex        =   38
         Top             =   900
         Width           =   1275
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cloth Description:"
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
         TabIndex        =   36
         Top             =   1755
         Width           =   2745
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cloth Name:"
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
         TabIndex        =   35
         Top             =   1335
         Width           =   2745
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity To Change:"
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
         TabIndex        =   33
         Top             =   2205
         Width           =   2745
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cost Price: (Rs.)"
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
         TabIndex        =   31
         Top             =   2655
         Width           =   2745
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cloth Type:"
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
         TabIndex        =   27
         Top             =   3060
         Width           =   2745
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cloth Type:"
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
         Left            =   1365
         TabIndex        =   19
         Top             =   3135
         Width           =   2745
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cost Price: (Rs.)"
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
         Left            =   1365
         TabIndex        =   18
         Top             =   2730
         Width           =   2745
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity:"
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
         Left            =   1365
         TabIndex        =   17
         Top             =   2280
         Width           =   2745
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cloth Serial No.:"
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
         Left            =   1380
         TabIndex        =   16
         Top             =   975
         Width           =   2745
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cloth Name:"
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
         Left            =   1365
         TabIndex        =   15
         Top             =   1410
         Width           =   2745
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cloth Description:"
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
         Left            =   1365
         TabIndex        =   14
         Top             =   1830
         Width           =   2745
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00764E10&
         Height          =   3465
         Left            =   75
         Top             =   390
         Width           =   8535
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
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00764E10&
      BorderWidth     =   2
      X1              =   5865
      X2              =   6285
      Y1              =   1245
      Y2              =   1065
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00764E10&
      BorderWidth     =   2
      X1              =   6255
      X2              =   8760
      Y1              =   1065
      Y2              =   1065
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00764E10&
      BorderWidth     =   2
      X1              =   30
      X2              =   8760
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
      Left            =   6180
      TabIndex        =   13
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
      TabIndex        =   12
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
      TabIndex        =   11
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
      TabIndex        =   10
      Top             =   375
      Width           =   7245
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   1470
      Left            =   15
      Picture         =   "purchase.frx":047A
      Stretch         =   -1  'True
      Top             =   30
      Width           =   2520
   End
End
Attribute VB_Name = "purchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbdescription_Click()
    cmbserial.ListIndex = cmbdescription.ListIndex
    cmbname.ListIndex = cmbdescription.ListIndex
    txtname2.Text = ""
    txtdescription2.Text = ""
    txtqty2.Text = ""
    txtcost2.Text = ""
    Option2(0).Value = True
    Option2(1).Value = False

End Sub

Private Sub cmbname_Click()
    cmbserial.ListIndex = cmbname.ListIndex
    cmbdescription.ListIndex = cmbname.ListIndex
    txtname2.Text = ""
    txtdescription2.Text = ""
    txtqty2.Text = ""
    txtcost2.Text = ""
    Option2(0).Value = True
    Option2(1).Value = False

End Sub

Private Sub cmbserial_Click()
    cmbname.ListIndex = cmbserial.ListIndex
    cmbdescription.ListIndex = cmbserial.ListIndex
    txtname2.Text = ""
    txtdescription2.Text = ""
    txtqty2.Text = ""
    txtcost2.Text = ""
    Option2(0).Value = True
    Option2(1).Value = False

End Sub

Private Sub cmdsave1_Click()
    
    If txtname.Text = "" Then
        MsgBox "Enter A Name Please:", vbCritical
        txtname.SetFocus
        Exit Sub
    End If
        
    If Not validity2(txtname, "Name") Then
        txtname.SetFocus
        Exit Sub
    End If
    
    If txtdescription.Text = "" Then
        MsgBox "Enter A Description Please:", vbCritical
        txtdescription.SetFocus
        Exit Sub
    End If
    
    If Not validity2(txtdescription, "Description") Then
        txtdescription.SetFocus
        Exit Sub
    End If
        
    If txtqty.Text = "" Then
        MsgBox "Enter A Quantity Please:", vbCritical
        txtqty.SetFocus
        Exit Sub
    End If
        
    If Not IsNumeric(txtqty.Text) Then
        MsgBox "Quantity should be Numeric!", vbCritical
        txtqty.SetFocus
        Exit Sub
    End If
    
    If CLng(txtqty.Text) < 0 Then
        MsgBox "Cannot Enter Negative Values In Quantity!", vbCritical
        txtqty.SetFocus
        Exit Sub
    End If
    
    If txtcost.Text = "" Then
        MsgBox "Enter A Cost Price Please:", vbCritical
        txtcost.SetFocus
        Exit Sub
    End If
        
    If Not IsNumeric(txtcost.Text) Then
        MsgBox "Cost Price should be Numeric!", vbCritical
        txtcost.SetFocus
        Exit Sub
    End If
    
    If CCur(txtcost.Text) < 0 Then
        MsgBox "Cannot Enter Negative Values In Cost!", vbCritical
        txtcost.SetFocus
        Exit Sub
    End If
    
    
    Call openconn
    sqlstr = "insert into clothes values('" & txtserial.Text & "','" & txtname.Text & "','" & txtdescription.Text & "'," & CLng(txtqty.Text) & "," & CCur(txtcost.Text) & "," & CBool(Option1(0).Value) & ")"
    Call rs(sqlstr)
    Call closeconn
    
    MsgBox "Cloth Entry Saved Successfully!", vbInformation
        
    Unload Me
    
End Sub


Private Sub cmdsave2_Click()
    If txtname2.Text = "" Then
        MsgBox "Enter A Name Please:", vbCritical
        txtname2.SetFocus
        Exit Sub
    End If
        
    If Not validity2(txtname2, "Name") Then
        txtname2.SetFocus
        Exit Sub
    End If
    
    If txtdescription2.Text = "" Then
        MsgBox "Enter A Description Please:", vbCritical
        txtdescription2.SetFocus
        Exit Sub
    End If
    
    If Not validity2(txtdescription2, "Description") Then
        txtdescription2.SetFocus
        Exit Sub
    End If
        
    If txtqty2.Text = "" Then
        MsgBox "Enter A Quantity Please:", vbCritical
        txtqty2.SetFocus
        Exit Sub
    End If
        
    If Not IsNumeric(txtqty2.Text) Then
        MsgBox "Quantity should be Numeric!", vbCritical
        txtqty2.SetFocus
        Exit Sub
    End If
    
    If CLng(txtqty2.Text) < 0 Then
        MsgBox "Cannot Enter Negative Values In Quantity!", vbCritical
        txtqty2.SetFocus
        Exit Sub
    End If
    
    If txtcost2.Text = "" Then
        MsgBox "Enter A Cost Price Please:", vbCritical
        txtcost2.SetFocus
        Exit Sub
    End If
        
    If Not IsNumeric(txtcost2.Text) Then
        MsgBox "Cost Price should be Numeric!", vbCritical
        txtcost2.SetFocus
        Exit Sub
    End If
    
    If CCur(txtcost2.Text) < 0 Then
        MsgBox "Cannot Enter Negative Values In Cost!", vbCritical
        txtcost2.SetFocus
        Exit Sub
    End If
        
    
    Call openconn
    sqlstr = "update clothes set name='" & txtname2.Text & "' where serialno = '" & cmbserial.Text & "'"
    Call rs(sqlstr)
    sqlstr = "update clothes set description='" & txtdescription2.Text & "' where serialno = '" & cmbserial.Text & "'"
    Call rs(sqlstr)
    sqlstr = "update clothes set qty=" & CLng(txtqty2.Text) & " where serialno = '" & cmbserial.Text & "'"
    Call rs(sqlstr)
    sqlstr = "update clothes set costprice=" & CCur(txtcost2.Text) & " where serialno = '" & cmbserial.Text & "'"
    Call rs(sqlstr)
    sqlstr = "update clothes set isstiched=" & CBool(Option2(0).Value) & " where serialno = '" & cmbserial.Text & "'"
    Call rs(sqlstr)
    
    Call closeconn
    
    MsgBox "Cloth Entry Saved Successfully!", vbInformation
        
    Unload Me

End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Command4_Click()
    If (cmbserial.ListIndex = -1) Then
        MsgBox "Please Select A Cloth from the Drop Down Box!", vbCritical
        cmbserial.SetFocus
    Exit Sub
    End If
    
    Call openconn
    sqlstr = "select qty, costprice, Isstiched from clothes where serialno='" & cmbserial.Text & "' and name = '" & cmbname.Text & "' and description = '" & cmbdescription.Text & "'"
    
    Call rs(sqlstr)
    
    If adoRS.EOF Then
        MsgBox "NO Cloth Selected", vbCritical
    Exit Sub
    End If
        
    txtname2.Text = cmbname.Text
    txtdescription2.Text = cmbdescription.Text
    txtqty2.Text = adoRS.Fields("qty")
    txtcost2.Text = adoRS.Fields("costprice")
    If adoRS.Fields("isstiched") Then
        Option2(0).Value = True
        Option2(1).Value = False
    Else
        Option2(0).Value = False
        Option2(1).Value = True
    End If
    Call closeconn
        
End Sub

Private Sub Form_Load()
    loginbar.Caption = "User = " & userid
    
    Call openconn
    sqlstr = "select serialno from clothes"
    Call rs(sqlstr)
  
      
    If (adoRS.EOF) Then
        txtserial.Text = "0000000001"
    Else
        hello = adoRS.Sort
        adoRS.MoveLast
        txtserial.Text = Right("0000000000" & CStr(CLng(adoRS.Fields(0)) + 1), 10)
    End If
    
    Call closeconn
    
    Call openconn
    sqlstr = "select serialno, name, description from clothes"
    Call rs(sqlstr)
    
    If Not (adoRS.EOF) Then
        hello = adoRS.Sort
        adoRS.MoveFirst
        While Not (adoRS.EOF)
            cmbserial.AddItem adoRS.Fields("serialno")
            cmbname.AddItem adoRS.Fields("name")
            cmbdescription.AddItem adoRS.Fields("description")
            adoRS.MoveNext
        Wend
    End If
    
    Call closeconn
    
    
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MainForm2.Show
End Sub

