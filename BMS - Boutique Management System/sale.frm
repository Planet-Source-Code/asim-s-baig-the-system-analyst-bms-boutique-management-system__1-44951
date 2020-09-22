VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form sale 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BMS - Clothes Selling"
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
   Icon            =   "sale.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   8820
   Begin TabDlg.SSTab SSTab1 
      Height          =   3945
      Left            =   45
      TabIndex        =   26
      Top             =   1365
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
      TabCaption(0)   =   "Clothes Selling"
      TabPicture(0)   =   "sale.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Shape1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label17"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Line5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label15"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label12"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label8"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label14"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label13"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label9"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label16"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label18"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "dt2"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmbdescription"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmbcloth"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmbserial"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Command1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cmdsave2"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtqty"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtqty2"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtprice"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "cmbaccount2"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "cmbname2"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtdescription"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Command4"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtname"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "isstiched"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtcost"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).ControlCount=   29
      TabCaption(1)   =   "Other Sellings And Income"
      TabPicture(1)   =   "sale.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txtamount"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Command3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmbname"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmbaccount"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "dt1"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label4"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label5"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Line6"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label10"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label11"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Shape2"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).ControlCount=   12
      Begin VB.TextBox txtcost 
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
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   2580
         Width           =   1335
      End
      Begin VB.TextBox isstiched 
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
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   3045
         Width           =   1515
      End
      Begin VB.TextBox txtname 
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
         Left            =   1380
         Locked          =   -1  'True
         MaxLength       =   49
         TabIndex        =   4
         Top             =   1230
         Width           =   3030
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
         Left            =   5565
         MaskColor       =   &H00987758&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1230
         UseMaskColor    =   -1  'True
         Width           =   2925
      End
      Begin VB.TextBox txtdescription 
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
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1680
         Width           =   3285
      End
      Begin VB.ComboBox cmbname2 
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
         Left            =   5790
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2115
         Width           =   2715
      End
      Begin VB.ComboBox cmbaccount2 
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
         Left            =   4365
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2115
         Width           =   1395
      End
      Begin VB.TextBox txtprice 
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
         Left            =   7320
         TabIndex        =   13
         Top             =   2985
         Width           =   1155
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
         Left            =   4905
         TabIndex        =   12
         Top             =   2985
         Width           =   1155
      End
      Begin VB.TextBox txtqty 
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
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2130
         Width           =   1155
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
         Left            =   6855
         MaskColor       =   &H00987758&
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3390
         UseMaskColor    =   -1  'True
         Width           =   1635
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
         Height          =   360
         Left            =   5235
         MaskColor       =   &H00987758&
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   3405
         UseMaskColor    =   -1  'True
         Width           =   1575
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
         Left            =   1395
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   810
         Width           =   1395
      End
      Begin VB.ComboBox cmbcloth 
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
         Left            =   2805
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   810
         Width           =   2715
      End
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
         Left            =   5550
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   810
         Width           =   2985
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
         Left            =   -69570
         MaskColor       =   &H00987758&
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2850
         UseMaskColor    =   -1  'True
         Width           =   1575
      End
      Begin VB.TextBox txtamount 
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
         Left            =   -71250
         MaxLength       =   99
         TabIndex        =   18
         Top             =   1740
         Width           =   2385
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E2D1D3&
         Caption         =   "Save Entry"
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
         Left            =   -71250
         MaskColor       =   &H00987758&
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2865
         UseMaskColor    =   -1  'True
         Width           =   1635
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
         Left            =   -69825
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1200
         Width           =   2715
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
         Left            =   -71250
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1200
         Width           =   1395
      End
      Begin MSComCtl2.DTPicker dt1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dddd, MMMM dd, yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   360
         Left            =   -71250
         TabIndex        =   19
         Top             =   2310
         Width           =   3810
         _ExtentX        =   6720
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
         CustomFormat    =   "dddd dd MMMMM, yyyy"
         Format          =   48824323
         CurrentDate     =   37657
      End
      Begin MSComCtl2.DTPicker dt2 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "dddd, MMMM dd, yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   360
         Left            =   4650
         TabIndex        =   11
         Top             =   2550
         Width           =   3810
         _ExtentX        =   6720
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
         CustomFormat    =   "dddd dd MMMMM, yyyy"
         Format          =   48824323
         CurrentDate     =   37657
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cost Price:"
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
         Left            =   315
         TabIndex        =   41
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Selling Date:"
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
         Left            =   1830
         TabIndex        =   40
         Top             =   2610
         Width           =   2745
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Type:"
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
         Left            =   330
         TabIndex        =   39
         Top             =   3105
         Width           =   975
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
         Left            =   60
         TabIndex        =   38
         Top             =   1305
         Width           =   1245
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
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
         Left            =   150
         TabIndex        =   37
         Top             =   1725
         Width           =   1155
      End
      Begin VB.Label Label8 
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
         Left            =   1560
         TabIndex        =   36
         Top             =   2175
         Width           =   2745
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Selling Price:"
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
         Left            =   5985
         TabIndex        =   35
         Top             =   3045
         Width           =   1305
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity To Sell:"
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
         Left            =   3225
         TabIndex        =   34
         Top             =   3045
         Width           =   1605
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Available:"
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
         Left            =   330
         TabIndex        =   33
         Top             =   2190
         Width           =   975
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
         Left            =   60
         TabIndex        =   32
         Top             =   885
         Width           =   1275
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00764E10&
         BorderWidth     =   2
         X1              =   120
         X2              =   8580
         Y1              =   750
         Y2              =   750
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Clothes Selling Entry"
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
         TabIndex        =   31
         Top             =   420
         Width           =   8445
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Date:"
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
         Left            =   -74070
         TabIndex        =   30
         Top             =   2370
         Width           =   2745
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Other Sellings And Income"
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
         Left            =   -74850
         TabIndex        =   29
         Top             =   405
         Width           =   8445
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00764E10&
         BorderWidth     =   2
         X1              =   -74880
         X2              =   -66420
         Y1              =   705
         Y2              =   705
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Amount Given By Customer: (Rs.)"
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
         Left            =   -74730
         TabIndex        =   28
         Top             =   1845
         Width           =   3375
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
         Left            =   -74055
         TabIndex        =   27
         Top             =   1260
         Width           =   2745
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
      X1              =   5850
      X2              =   6270
      Y1              =   1230
      Y2              =   1050
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00764E10&
      BorderWidth     =   2
      X1              =   6240
      X2              =   8745
      Y1              =   1050
      Y2              =   1050
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00764E10&
      BorderWidth     =   2
      X1              =   15
      X2              =   8745
      Y1              =   1260
      Y2              =   1260
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
      TabIndex        =   25
      Top             =   1035
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
      TabIndex        =   24
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
      TabIndex        =   23
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
      TabIndex        =   22
      Top             =   375
      Width           =   7245
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   1470
      Left            =   15
      Picture         =   "sale.frx":047A
      Stretch         =   -1  'True
      Top             =   30
      Width           =   2520
   End
End
Attribute VB_Name = "sale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim detail_got As Boolean

Private Sub cmbaccount2_Click()
  cmbname2.ListIndex = cmbaccount2.ListIndex
End Sub

Private Sub cmbname2_Click()
    cmbaccount2.ListIndex = cmbname2.ListIndex
    
End Sub

Private Sub cmbdescription_Click()
    cmbserial.ListIndex = cmbdescription.ListIndex
    cmbcloth.ListIndex = cmbdescription.ListIndex
    txtname.Text = ""
    txtdescription.Text = ""
    txtqty.Text = ""
    txtcost.Text = ""
    isstiched.Text = ""
    
End Sub

Private Sub cmbcloth_Click()
    cmbserial.ListIndex = cmbcloth.ListIndex
    cmbdescription.ListIndex = cmbcloth.ListIndex
    txtname.Text = ""
    txtdescription.Text = ""
    txtqty.Text = ""
    txtcost.Text = ""
    isstiched.Text = ""
    
End Sub

Private Sub cmbserial_Click()
    cmbcloth.ListIndex = cmbserial.ListIndex
    cmbdescription.ListIndex = cmbserial.ListIndex
    txtname.Text = ""
    txtdescription.Text = ""
    txtqty.Text = ""
    txtcost.Text = ""
    isstiched.Text = ""
    
End Sub


Private Sub cmbaccount_Click()
    cmbname.ListIndex = cmbaccount.ListIndex
    txtamount.Text = ""
End Sub

Private Sub cmbname_Click()
    cmbaccount.ListIndex = cmbname.ListIndex
    txtamount.Text = ""
End Sub

Private Sub cmdsave2_Click()
    If (cmbserial.ListIndex = -1) Then
        MsgBox "Please Select A Cloth To Sell!", vbCritical
        cmbserial.SetFocus
        Exit Sub
    End If
    
    If detail_got = False Then
        MsgBox "Please Click The Get Details Button!", vbCritical
        Exit Sub
    End If
    
    If (cmbaccount2.ListIndex = -1) Then
        MsgBox "Please Select A Customer!", vbCritical
        cmbaccount2.SetFocus
        Exit Sub
    End If
    
    If txtqty2.Text = "" Then
        MsgBox "Please Enter Quantity To Sell or Enter 0 in it!", vbCritical
        txtqty2.SetFocus
        Exit Sub
    End If
        
    If Not IsNumeric(txtqty2.Text) Then
        MsgBox "Quantity to Sell Has to be Number Format!", vbCritical
        txtqty2.SetFocus
        Exit Sub
    End If
    
    If CLng(txtqty2.Text) < 0 Then
        MsgBox "Cannot Enter Negative Values In Quantity!", vbCritical
        txtqty2.SetFocus
        Exit Sub
    End If
    
    If (CCur(txtqty.Text) < CCur(txtqty2.Text)) Then
        testing = MsgBox("You Are Selling more in Quantity than Available!" & vbCrLf & vbCrLf & "PLEASE NOTE THAT YOU ARE JUST SELLING MORE QUANTITY THAN AVAILABLE IN STOCK!" & vbCrLf & vbCrLf & "Do You Really Want To Sell More Than Available?", vbYesNo + vbCritical)
        If testing = vbYes Then
        Else
            txtqty2.SetFocus
            Exit Sub
        End If
    End If
    
    
    If txtprice.Text = "" Then
        MsgBox "Please Enter Selling Price!", vbCritical
        txtprice.SetFocus
        Exit Sub
    End If
        
    If Not IsNumeric(txtprice.Text) Then
        MsgBox "Selling Price Has to be Number Format!", vbCritical
        txtprice.SetFocus
        Exit Sub
    End If
    
    If CCur(txtprice.Text) < 0 Then
        MsgBox "Cannot Enter Negative Values In Price!", vbCritical
        txtprice.SetFocus
        Exit Sub
    End If
    
    If (CCur(txtprice.Text) < CCur(txtcost.Text)) Then
        testing = MsgBox("You Are Selling The Cloth at Lesser Price Than Its Cost, This is a LOSS!" & vbCrLf & vbCrLf & "Do You Want To Proceed?", vbYesNo + vbCritical)
        If testing = vbYes Then
        Else
            txtprice.SetFocus
            Exit Sub
        End If
    End If
    
    
    dated = dt2.Value
    dated = CDate(dated)
    
    Call openconn
    
    sqlstr = "insert into selling values('" & cmbserial.Text & "','" & cmbaccount2.Text & "'," & CLng(txtqty2.Text) & "," & CCur(txtprice.Text) & ",#" & dated & "#)"
    
    Call rs(sqlstr)
    
    Call closeconn
    
    
    Call openconn
    
    sqlstr = "update clothes set qty = " & CLng(CLng(txtqty.Text) - CLng(txtqty2.Text)) & " where serialno = '" & cmbserial.Text & "'"
    
    Call rs(sqlstr)
    
    Call closeconn
    
    msg = "Amount Given By Customer Saved Successfully!"
    
    If (CLng(txtqty.Text) < CLng(txtqty2.Text)) Then
        check = True
        msg = msg & vbCrLf & vbCrLf & "PLEASE NOTE THAT YOU HAVE JUST SOLD MORE QUANTITY THAN AVAILABLE IN STOCK!"
    End If
    
    If check Then
        MsgBox msg, vbExclamation
    Else
        MsgBox msg, vbInformation
    End If
        
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
    
    If txtamount.Text = "" Then
        MsgBox "Please Enter Amount or Enter 0 in it!", vbCritical
        txtamount.SetFocus
        Exit Sub
    End If
        
    If Not validity2(txtamount, "Amount Given By Customer") Then
        txtamount.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(txtamount.Text) Then
        MsgBox "Customer Amount Has to be Number Format!", vbCritical
        txtamount.SetFocus
    Exit Sub
    End If
        
       
    If CCur(txtamount.Text) < 0 Then
        MsgBox "Cannot Enter Negative Values In Amount!", vbCritical
        txtamount.SetFocus
        Exit Sub
    End If
            
    dated = dt1.Value
    dated = CDate(dated)
       
    Call openconn
    
    sqlstr = "insert into income values('" & cmbaccount.Text & "'," & CCur(txtamount.Text) & ",#" & dated & "#)"
    
    Call rs(sqlstr)
    
    Call closeconn
    
    MsgBox "Amount Given By Customer Saved Successfully!", vbInformation
        
    Unload Me

End Sub

Private Sub Command4_Click()
    If (cmbserial.ListIndex = -1) Then
        MsgBox "Please Select A Cloth from the Drop Down Box!", vbCritical
        cmbserial.SetFocus
    Exit Sub
    End If
    
    Call openconn
    sqlstr = "select qty, costprice, Isstiched from clothes where serialno='" & cmbserial.Text & "' and name = '" & cmbcloth.Text & "' and description = '" & cmbdescription.Text & "'"
    
    Call rs(sqlstr)
    
    If adoRS.EOF Then
        MsgBox "NO Cloth Selected", vbCritical
        Call closeconn
    Exit Sub
    End If
        
    txtname.Text = cmbcloth.Text
    txtdescription.Text = cmbdescription.Text
    txtcost.Text = adoRS.Fields("costprice")
    txtqty.Text = adoRS.Fields("qty")
    If adoRS.Fields("isstiched") Then
        isstiched.Text = "Stiched"
    Else
        isstiched.Text = "Un-Stiched"
    End If
    Call closeconn
        
    detail_got = True

End Sub

Private Sub Form_Load()
    loginbar.Caption = "User = " & userid
    
    detail_got = False
    
    dt1.Value = Date
    dt2.Value = Date
    
    Call openconn
    sqlstr = "select serialno, name, description from clothes"
    Call rs(sqlstr)
    
    If Not (adoRS.EOF) Then
        hello = adoRS.Sort
        adoRS.MoveFirst
        While Not (adoRS.EOF)
            cmbserial.AddItem adoRS.Fields("serialno")
            cmbcloth.AddItem adoRS.Fields("name")
            cmbdescription.AddItem adoRS.Fields("description")
            adoRS.MoveNext
        Wend
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
            cmbaccount2.AddItem adoRS.Fields("customerid")
            cmbname.AddItem adoRS.Fields("name")
            cmbname2.AddItem adoRS.Fields("name")
            adoRS.MoveNext
        Wend
    End If
    
    Call closeconn

    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MainForm2.Show
End Sub

