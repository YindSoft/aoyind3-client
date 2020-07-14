VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   ".::AoYind::."
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   555
   ClientWidth     =   15360
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMainN.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox SendTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   360
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   2640
      Visible         =   0   'False
      Width           =   8040
   End
   Begin VB.TextBox SendCMSTXT 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   360
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   2640
      Visible         =   0   'False
      Width           =   8040
   End
   Begin VB.PictureBox pRender 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2760
      Left            =   195
      ScaleHeight     =   184
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   736
      TabIndex        =   7
      Top             =   2580
      Width           =   11040
      Begin VB.TextBox tPass 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   6420
         MaxLength       =   160
         PasswordChar    =   "*"
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Chat"
         Top             =   7260
         Visible         =   0   'False
         Width           =   2700
      End
      Begin VB.TextBox tUser 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   6420
         MaxLength       =   160
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Chat"
         Top             =   9285
         Visible         =   0   'False
         Width           =   2700
      End
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3975
      Left            =   11280
      ScaleHeight     =   265
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   219
      TabIndex        =   16
      Top             =   2790
      Width           =   3285
      Begin VB.Label lblItem 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   600
         TabIndex        =   17
         Top             =   120
         Width           =   2175
      End
   End
   Begin VB.PictureBox picHechiz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      DrawStyle       =   3  'Dash-Dot
      ForeColor       =   &H00FFFFFF&
      Height          =   2625
      Left            =   12360
      Picture         =   "frmMainN.frx":7265
      ScaleHeight     =   175
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   140
      TabIndex        =   15
      Top             =   3240
      Width           =   2100
   End
   Begin VB.PictureBox BarraHechiz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2595
      Left            =   14760
      Picture         =   "frmMainN.frx":191C3
      ScaleHeight     =   173
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   13
      Top             =   3360
      Width           =   240
      Begin VB.PictureBox BarritaHechiz 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   105
         Left            =   15
         Picture         =   "frmMainN.frx":1B275
         ScaleHeight     =   7
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   14
         TabIndex        =   14
         Top             =   240
         Width           =   210
      End
   End
   Begin VB.Timer macrotrabajo 
      Enabled         =   0   'False
      Left            =   9000
      Top             =   1320
   End
   Begin VB.Timer Macro 
      Interval        =   750
      Left            =   5760
      Top             =   1440
   End
   Begin VB.Timer Second 
      Enabled         =   0   'False
      Interval        =   1050
      Left            =   5760
      Top             =   960
   End
   Begin VB.Timer SpoofCheck 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   8520
      Top             =   1320
   End
   Begin VB.PictureBox pConsola 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1275
      Left            =   480
      Picture         =   "frmMainN.frx":1B3EB
      ScaleHeight     =   85
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   522
      TabIndex        =   4
      Top             =   720
      Width           =   7830
      Begin MSWinsockLib.Winsock WSock 
         Left            =   3240
         Top             =   720
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Timer tMouse 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   7080
         Top             =   0
      End
   End
   Begin VB.PictureBox BarraConsola 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1365
      Left            =   10920
      Picture         =   "frmMainN.frx":1D3EF
      ScaleHeight     =   91
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   2
      Top             =   840
      Width           =   270
      Begin VB.PictureBox BarritaConsola 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   105
         Left            =   30
         Picture         =   "frmMainN.frx":1E819
         ScaleHeight     =   7
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   14
         TabIndex        =   3
         Top             =   1020
         Width           =   210
      End
   End
   Begin VB.Label lblEXP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   12480
      TabIndex        =   31
      Top             =   1680
      Width           =   1920
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   14400
      Top             =   0
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   3
      Left            =   13680
      MouseIcon       =   "frmMainN.frx":1E98F
      MousePointer    =   99  'Custom
      Top             =   7770
      Width           =   1410
   End
   Begin VB.Label lblAgilidad 
      BackStyle       =   0  'Transparent
      Caption         =   "A: 18"
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   9120
      TabIndex        =   30
      Top             =   11160
      Width           =   690
   End
   Begin VB.Label lblFuerza 
      BackStyle       =   0  'Transparent
      Caption         =   "F: 20"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   9120
      TabIndex        =   29
      Top             =   10920
      Width           =   570
   End
   Begin VB.Label lblUsers 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   12240
      TabIndex        =   28
      Top             =   10200
      Width           =   690
   End
   Begin VB.Label lblSed 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   11790
      TabIndex        =   27
      Top             =   9375
      Width           =   1440
   End
   Begin VB.Label lblHambre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   11790
      TabIndex        =   26
      Top             =   9030
      Width           =   1440
   End
   Begin VB.Label lblVida 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   11790
      TabIndex        =   25
      Top             =   8700
      Width           =   1440
   End
   Begin VB.Label lblMana 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   11790
      TabIndex        =   24
      Top             =   8265
      Width           =   1440
   End
   Begin VB.Label lblSedN 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   11760
      TabIndex        =   23
      Top             =   9360
      Width           =   1455
   End
   Begin VB.Label lblHambreN 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   11760
      TabIndex        =   22
      Top             =   9015
      Width           =   1455
   End
   Begin VB.Label lblVidaN 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   11760
      TabIndex        =   21
      Top             =   8685
      Width           =   1455
   End
   Begin VB.Label lblManaN 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   11760
      TabIndex        =   20
      Top             =   8250
      Width           =   1455
   End
   Begin VB.Label lblEnergia 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   11790
      TabIndex        =   19
      Top             =   7920
      Width           =   1440
   End
   Begin VB.Label lblEnergiaN 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   11760
      TabIndex        =   18
      Top             =   7905
      Width           =   1455
   End
   Begin VB.Image btnHechizos 
      Height          =   750
      Left            =   13560
      MousePointer    =   99  'Custom
      Picture         =   "frmMainN.frx":1EAE1
      Top             =   2280
      Width           =   1530
   End
   Begin VB.Image btnInventario 
      Height          =   750
      Left            =   11760
      MousePointer    =   99  'Custom
      Picture         =   "frmMainN.frx":23F3E
      Top             =   2280
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Label Coord 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "(000,000)"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   13560
      TabIndex        =   12
      Top             =   10440
      Width           =   825
   End
   Begin VB.Image PicSeg 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   7800
      Picture         =   "frmMainN.frx":29378
      Stretch         =   -1  'True
      Top             =   11040
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   14880
      Top             =   0
      Width           =   255
   End
   Begin VB.Label LvlLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   12600
      TabIndex        =   11
      Top             =   1080
      Width           =   120
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   375
      Index           =   0
      Left            =   11640
      MouseIcon       =   "frmMainN.frx":29830
      MousePointer    =   99  'Custom
      Top             =   3060
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   375
      Index           =   1
      Left            =   11640
      MouseIcon       =   "frmMainN.frx":29982
      MousePointer    =   99  'Custom
      Top             =   2640
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CmdLanzar 
      Height          =   525
      Left            =   11760
      MouseIcon       =   "frmMainN.frx":29AD4
      MousePointer    =   99  'Custom
      Top             =   6720
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.Image cmdInfo 
      Height          =   525
      Left            =   14040
      MouseIcon       =   "frmMainN.frx":29C26
      MousePointer    =   99  'Custom
      Top             =   6720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label GldLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10000000"
      ForeColor       =   &H0000C0C0&
      Height          =   195
      Left            =   10320
      TabIndex        =   10
      Top             =   11040
      Width           =   840
   End
   Begin VB.Image Image3 
      Height          =   315
      Index           =   0
      Left            =   9960
      Top             =   10920
      Width           =   390
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   0
      Left            =   13680
      MouseIcon       =   "frmMainN.frx":29D78
      MousePointer    =   99  'Custom
      Top             =   9360
      Width           =   1410
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   1
      Left            =   13680
      MouseIcon       =   "frmMainN.frx":29ECA
      MousePointer    =   99  'Custom
      Top             =   9000
      Width           =   1410
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   2
      Left            =   13680
      MouseIcon       =   "frmMainN.frx":2A01C
      MousePointer    =   99  'Custom
      Top             =   8520
      Width           =   1410
   End
   Begin VB.Label lbCRIATURA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   5.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   120
      Left            =   9000
      TabIndex        =   9
      Top             =   2400
      Width           =   30
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "El Yind"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   12480
      TabIndex        =   8
      Top             =   840
      Width           =   2145
   End
   Begin VB.Image STAShp 
      Height          =   180
      Left            =   11790
      Picture         =   "frmMainN.frx":2A16E
      Top             =   7905
      Width           =   1395
   End
   Begin VB.Image MANShp 
      Height          =   180
      Left            =   11790
      Picture         =   "frmMainN.frx":2AA82
      Top             =   8250
      Width           =   1395
   End
   Begin VB.Image Hpshp 
      Height          =   180
      Left            =   11790
      Picture         =   "frmMainN.frx":2B396
      Top             =   8685
      Width           =   1395
   End
   Begin VB.Image COMIDAsp 
      Height          =   180
      Left            =   11790
      Picture         =   "frmMainN.frx":2BCAA
      Top             =   9015
      Width           =   1395
   End
   Begin VB.Image AGUAsp 
      Height          =   180
      Left            =   11790
      Picture         =   "frmMainN.frx":2C5BE
      Top             =   9360
      Width           =   1395
   End
   Begin VB.Label lblExpN 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   12480
      TabIndex        =   32
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Image iBEXP 
      Height          =   135
      Left            =   12360
      Picture         =   "frmMainN.frx":2CED2
      Top             =   1680
      Width           =   2025
   End
   Begin VB.Image iBEXPE 
      Height          =   135
      Left            =   11160
      Picture         =   "frmMainN.frx":2DD6E
      Top             =   1485
      Width           =   60
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'AoYind 3.0
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Public tX As Integer
Public tY As Integer
Public MouseX As Integer
Public MouseY As Integer
Public MouseBoton As Long
Public MouseShift As Long
Private clicX As Long
Private clicY As Long
Public SinOrtografia As Boolean
'Dim gDSB As DirectSoundBuffer
'Dim gD As DSBUFFERDESC
'Dim gW As WAVEFORMATEX
Dim gFileName As String
'Dim dsE As DirectSoundEnum
'Dim Pos(0) As DSBPOSITIONNOTIFY
Public IsPlaying As PlayLoop

Dim PuedeMacrear As Boolean
Dim OldYConsola As Integer
Public hlst As clsGraphicalList
Dim InvX As Integer
Dim InvY As Integer
Public WithEvents Client As CSocketMaster
Attribute Client.VB_VarHelpID = -1

Private Sub Client_Connect()
    'Clean input and output buffers
    Call incomingData.ReadASCIIStringFixed(incomingData.Length)
    Call outgoingData.ReadASCIIStringFixed(outgoingData.Length)
    
#If SeguridadAlkon Then
    Call ConnectionStablished(Socket1.PeerAddress)
#End If
    
    Second.Enabled = True

    Select Case EstadoLogin
        Case E_MODO.CrearNuevoPj
#If SeguridadAlkon Then
            Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
#End If
            Call Login
        
        Case E_MODO.Normal
#If SeguridadAlkon Then
            Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
#End If
            Call Login
            iServer = 0
            iCliente = 0
            DummyCode = StrConv("damn" & StrReverse(UCase$(UserName)) & "you", vbFromUnicode)

        Case E_MODO.Cuentas
#If SeguridadAlkon Then
            Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
#End If
            Call Login
    End Select
End Sub

Private Sub Client_CloseSck()

    Dim i As Long
    Client.CloseSck
    
Call ClosePj
End Sub

Private Sub Client_DataArrival(ByVal bytesTotal As Long)
    Dim RD As String
    Dim Data() As Byte
    
    Client.GetData RD
    Data = StrConv(RD, vbFromUnicode)
    
    Call DataCorrect(DummyCode, Data, iServer)
    
    'Set data in the buffer
    Call incomingData.WriteBlock(Data)
    
    NotEnoughData = False
    
    'Send buffer to Handle data
    Call HandleIncomingData
End Sub

Private Sub Client_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    '*********************************************
    'Handle socket errors
    '*********************************************
    If Number = 24036 Then
        Call MessageBox("Por favor espere, intentando completar conexion.", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
        Exit Sub
    ElseIf Number = 10049 Then
        Call MessageBox("Su equipo no soporta la API de Socket, se cambiará su configuración a Winsock, si problema persiste contacte soporte.", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
    End If
    
    Call MessageBox(Description, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")

    Second.Enabled = False

    Client.CloseSck
    

    If Not frmCrearPersonaje.Visible And Not Conectar Then
        Call ClosePj
    Else
        frmCrearPersonaje.MousePointer = 0
    End If

End Sub

Private Sub BarraConsola_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim TempY As Integer
    Dim TamCon As Integer
    TempY = Y - 3
    TamCon = (LineasConsola - 6)
    If TamCon > 0 Then
        If TempY < 16 Then
            If OffSetConsola > 0 Then OffSetConsola = OffSetConsola - 1
            TempY = 16 + OffSetConsola * 52 / TamCon
        ElseIf TempY > 68 Then
            If OffSetConsola < TamCon Then OffSetConsola = OffSetConsola + 1
            TempY = 16 + OffSetConsola * 52 / TamCon
        Else
            If LineasConsola <= 6 Then TempY = 68
            OffSetConsola = Int((TempY - 16) * TamCon / 52)
        End If
    Else
        TempY = 68
    End If
    BarritaConsola.Top = TempY
    ReDrawConsola
End Sub

Private Sub BarraHechiz_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim TempY As Integer
    Dim TamCon As Integer
    TempY = Y - 3
    Dim MaxItems As Integer
    MaxItems = Int(picHechiz.Height / hlst.Pixel_Alto)
    TamCon = (hlst.ListCount - MaxItems)
    
    If TamCon > 0 Then
        If TempY < 16 Then
            If hlst.Scroll > 0 Then hlst.Scroll = hlst.Scroll - 1
            TempY = 16 + hlst.Scroll * 134 / TamCon
        ElseIf TempY > 150 Then
            If hlst.Scroll < TamCon Then hlst.Scroll = hlst.Scroll + 1
            TempY = 16 + hlst.Scroll * 134 / TamCon
        Else
            If hlst.ListCount <= MaxItems Then TempY = 150
            hlst.Scroll = Int((TempY - 16) * TamCon / 134)
        End If
    Else
        TempY = 150
    End If
    BarritaHechiz.Top = TempY
End Sub

Private Sub BarritaConsola_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 Then
    OldYConsola = Y
End If
End Sub

Private Sub BarritaConsola_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 Then
    Dim TempY As Integer
    TempY = BarritaConsola.Top + (Y - OldYConsola)
    If TempY < 16 Then TempY = 16
    If TempY > 68 Then TempY = 68
    If LineasConsola <= 6 Then TempY = 68
    OffSetConsola = Int((TempY - 16) * (LineasConsola - 6) / 52)
    BarritaConsola.Top = TempY
    ReDrawConsola
End If
End Sub

Private Sub BarritaHechiz_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 Then
    hlst.OldY = Y
End If
End Sub

Private Sub BarritaHechiz_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 Then
    Dim TempY As Integer
    Dim MaxItems As Integer
    MaxItems = Int(picHechiz.Height / hlst.Pixel_Alto)
    TempY = BarritaHechiz.Top + (Y - hlst.OldY)
    If TempY < 16 Then TempY = 16
    If TempY > 150 Then TempY = 150
    If hlst.ListCount <= MaxItems Then TempY = 150
    hlst.Scroll = Int((TempY - 16) * (hlst.ListCount - MaxItems) / 134)
    BarritaHechiz.Top = TempY
End If
End Sub

Private Sub btnHechizos_Click()
    Call Audio.Sound_Play(SND_CLICK)
    picInv.Visible = False
    hlst.Visible = True
    cmdInfo.Visible = True
    CmdLanzar.Visible = True
    
    cmdMoverHechi(0).Visible = True
    cmdMoverHechi(1).Visible = True
    
    cmdMoverHechi(0).Enabled = True
    cmdMoverHechi(1).Enabled = True
    
    btnInventario.Visible = True
    btnHechizos.Visible = False
End Sub

Private Sub btnInventario_Click()
    Call Audio.Sound_Play(SND_CLICK)
    picInv.Visible = True

    hlst.Visible = False
    cmdInfo.Visible = False
    CmdLanzar.Visible = False
    
    cmdMoverHechi(0).Visible = True
    cmdMoverHechi(1).Visible = True
    
    cmdMoverHechi(0).Enabled = False
    cmdMoverHechi(1).Enabled = False
    
    btnInventario.Visible = False
    btnHechizos.Visible = True
End Sub

Private Sub cmdMoverHechi_Click(Index As Integer)
    If hlst.ListIndex = -1 Then Exit Sub
    Dim sTemp As String

    Select Case Index
        Case 1 'subir
            If hlst.ListIndex = 0 Then Exit Sub
        Case 0 'bajar
            If hlst.ListIndex = hlst.ListCount - 1 Then Exit Sub
    End Select

    Call WriteMoveSpell(Index, hlst.ListIndex + 1)
    
    Select Case Index
        Case 1 'subir
            sTemp = hlst.List(hlst.ListIndex - 1)
            hlst.List(hlst.ListIndex - 1) = hlst.List(hlst.ListIndex)
            hlst.List(hlst.ListIndex) = sTemp
            hlst.ListIndex = hlst.ListIndex - 1
        Case 0 'bajar
            sTemp = hlst.List(hlst.ListIndex + 1)
            hlst.List(hlst.ListIndex + 1) = hlst.List(hlst.ListIndex)
            hlst.List(hlst.ListIndex) = sTemp
            hlst.ListIndex = hlst.ListIndex + 1
    End Select
End Sub

Public Sub DibujarSeguro()
PicSeg.Visible = True
End Sub

Public Sub DesDibujarSeguro()
PicSeg.Visible = False
End Sub






Private Sub Command1_Click()
'Dim i As Integer
'i = Val(Text1.Text)
'If i > 0 Then
'Call DrawTextPergamino(Tutoriales(i).Linea1 & vbCrLf & Tutoriales(i).Linea2 & vbCrLf & Tutoriales(i).Linea3, 0, 0)
'End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case CustomKeys.BindedKey(eKeyType.mKeyVerMapa)
        If (Not SendTxt.Visible) And (Not SendCMSTXT.Visible) Then
            VerMapa = True
        End If
    Case vbKeyEscape
        If Conectar Then
            If GTCPres < 10000 Then
                GTCInicial = GTCInicial - (10000 - GTCPres)
                Call Audio.MusicMP3Play("10.mp3")
            ElseIf MostrarEntrar > 0 Then
                If GTCPres - MostrarEntrar > 1000 Then
                    MostrarEntrar = -GTCPres
                    tUser.Visible = False
                    tPass.Visible = False
                    Call Audio.Sound_Play(SND_CADENAS)
                End If
            Else
                prgRun = False
            End If
        End If
End Select
End Sub
Public Sub SetRender(Full As Boolean)
If Full Then
    pRender.Move 0, 0, 1024, 768
Else
    pRender.Move 13, 169, 736, 544
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If Conectar Then Exit Sub

#If SeguridadAlkon Then
    If LOGGING Then Call CheatingDeath.StoreKey(KeyCode, False)
#End If
    
    If (Not SendTxt.Visible) And (Not SendCMSTXT.Visible) Then
        
        'Checks if the key is valid
        If LenB(CustomKeys.ReadableName(KeyCode)) > 0 Then
            Select Case KeyCode
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleMusic)
                    Audio.MusicActivated = Not Audio.MusicActivated
                
                Case CustomKeys.BindedKey(eKeyType.mKeyGetObject)
                    Call AgarrarItem
                Case CustomKeys.BindedKey(eKeyType.mKeyEquipObject)
                    Call EquiparItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleNames)
                    Nombres = Not Nombres
                
                Case CustomKeys.BindedKey(eKeyType.mKeyTamAnimal)
                    If UserEstado = 1 Then
                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                        End With
                    Else
                        Call WriteWork(eSkill.Domar)
                    End If
                    
                Case CustomKeys.BindedKey(eKeyType.mKeySteal)
                    If UserEstado = 1 Then
                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                        End With
                    Else
                        Call WriteWork(eSkill.Robar)
                    End If
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyHide)
                    If UserEstado = 1 Then
                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                        End With
                    Else
                        If MainTimer.Check(TimersIndex.Hide) Then
                            Call WriteWork(eSkill.Ocultarse)
                        End If
                    End If
                                    
                Case CustomKeys.BindedKey(eKeyType.mKeyDropObject)
                    Call TirarItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyUseObject)
                    If macrotrabajo.Enabled Then DesactivarMacroTrabajo
                        
                    If MainTimer.Check(TimersIndex.UseItemWithU) Then
                        Call UsarItem
                    End If
                
                Case CustomKeys.BindedKey(eKeyType.mKeyRequestRefresh)
                    If MainTimer.Check(TimersIndex.SendRPU) And Not UserEmbarcado Then
                        Call WriteRequestPositionUpdate
                        Beep
                    End If
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleResuscitationSafe)
                    Call WriteResuscitationToggle
            End Select
        Else
            Select Case KeyCode
                'Custom messages!
                Case vbKey0 To vbKey9
                    If LenB(CustomMessages.Message((KeyCode - 39) Mod 10)) <> 0 Then
                        Call WriteTalk(CustomMessages.Message((KeyCode - 39) Mod 10))
                    End If
            End Select
        End If
    End If
    
    Select Case KeyCode
        Case CustomKeys.BindedKey(eKeyType.mKeyTalkWithGuild)
            If SendTxt.Visible Then Exit Sub
            
            If (Not frmComerciar.Visible) And (Not frmComerciarUsu.Visible) And _
              (Not frmBancoObj.Visible) And _
              (Not frmMSG.Visible) And _
              (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) Then
                SendCMSTXT.Visible = True
                SendCMSTXT.SetFocus
            End If
        Case CustomKeys.BindedKey(eKeyType.mKeyVerMapa)
            VerMapa = False
        Case CustomKeys.BindedKey(eKeyType.mKeyTakeScreenShot)
            Call ScreenCapture
        
        Case CustomKeys.BindedKey(eKeyType.mKeyToggleFPS)
            FPSFLAG = Not FPSFLAG
            
        Case CustomKeys.BindedKey(eKeyType.mKeyShowOptions)
            Call frmOpciones.Show(vbModeless, frmMain)
        
        Case CustomKeys.BindedKey(eKeyType.mKeyMeditate)
            If UserMinMAN = UserMaxMAN Then Exit Sub
            
            If UserEstado = 1 Then
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                End With
                Exit Sub
            End If
                
            If Not PuedeMacrear Then
                AddtoRichPicture "¡No puedes usar el macro tan rápido!", 255, 255, 255, True, False, False
            ElseIf charlist(UserCharIndex).Moving = 0 Then
                Call WriteMeditate
                PuedeMacrear = False
            End If

        Case CustomKeys.BindedKey(eKeyType.mKeyWorkMacro)
            If UserEstado = 1 Then
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                End With
                Exit Sub
            End If
            
            If macrotrabajo.Enabled Then
                DesactivarMacroTrabajo
            Else
                ActivarMacroTrabajo
            End If
        
        Case CustomKeys.BindedKey(eKeyType.mKeyExitGame)
            If frmMain.macrotrabajo.Enabled Then DesactivarMacroTrabajo
            Call WriteQuit
            
        Case CustomKeys.BindedKey(eKeyType.mKeyAttack)
            If Shift <> 0 Then Exit Sub
            
            If Not MainTimer.Check(TimersIndex.Arrows, False) Then Exit Sub 'Check if arrows interval has finished.
            If Not MainTimer.Check(TimersIndex.CastSpell, False) Then 'Check if spells interval has finished.
                If Not MainTimer.Check(TimersIndex.CastAttack) Then Exit Sub 'Corto intervalo Golpe-Hechizo
            Else
                If Not MainTimer.Check(TimersIndex.Attack) Or UserDescansar Or UserMeditar Then Exit Sub
            End If
            
            If macrotrabajo.Enabled Then DesactivarMacroTrabajo
            Call WriteAttack
        
        Case CustomKeys.BindedKey(eKeyType.mKeyTalk)
            If SendCMSTXT.Visible Then Exit Sub
            
            If (Not frmComerciar.Visible) And (Not frmComerciarUsu.Visible) And _
              (Not frmBancoObj.Visible) And _
              (Not frmMSG.Visible) And _
              (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) Then
                SendTxt.Visible = True
                SendTxt.SetFocus
            End If
            
    End Select
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Y < 24 And NoRes Then
    MoverVentana (Me.hwnd)
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If UserPasarNivel > 0 Then
    frmMain.lblEXP.Caption = Round((UserExp / UserPasarNivel) * 100, 2) & "%"
    frmMain.lblExpN.Caption = Round((UserExp / UserPasarNivel) * 100, 2) & "%"
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If prgRun = True Then
        prgRun = False
        Cancel = 1
    End If
End Sub

Private Sub Image5_Click()

End Sub

Private Sub Image2_Click()
prgRun = False
End Sub


Private Sub Label3_Click()

End Sub

Private Sub Image4_Click()
Me.WindowState = vbMinimized
End Sub

Private Sub lblEXP_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    frmMain.lblEXP.Caption = UserExp & "/" & UserPasarNivel
    frmMain.lblExpN.Caption = UserExp & "/" & UserPasarNivel
End Sub

Private Sub Macro_Timer()
    PuedeMacrear = True
End Sub

Private Sub macrotrabajo_Timer()
    If Inventario.SelectedItem = 0 Then
        DesactivarMacroTrabajo
        Exit Sub
    End If
    
    'Macros are disabled if not using Argentum!
    'If Not Api.IsAppActive() Then
    '    DesactivarMacroTrabajo
    '    Exit Sub
    'End If
    
    If (UsingSkill = eSkill.Pesca Or UsingSkill = eSkill.Talar Or UsingSkill = eSkill.Mineria Or UsingSkill = FundirMetal Or UsingSkill = eSkill.Herreria) Then
        Call WriteWorkLeftClick(tX, tY, UsingSkill)
        UsingSkill = 0
    End If
    
    'If Inventario.OBJType(Inventario.SelectedItem) = eObjType.otWeapon Then
     Call UsarItem
End Sub
Public Sub ControlSeguroResu(ByVal Mostrar As Boolean)
If Mostrar Then
    'If Not PicResu.Visible Then
    '    PicResu.Visible = True
    'End If
Else
    'If PicResu.Visible Then
    '    PicResu.Visible = False
    'End If
End If
End Sub
Public Sub ActivarMacroTrabajo()
    macrotrabajo.Interval = INT_MACRO_TRABAJO
    macrotrabajo.Enabled = True
    Call AddtoRichPicture("Macro Trabajo ACTIVADO", 0, 200, 200, False, True, False)
End Sub

Public Sub DesactivarMacroTrabajo()
    macrotrabajo.Enabled = False
    MacroBltIndex = 0
    Call AddtoRichPicture("Macro Trabajo DESACTIVADO", 0, 200, 200, False, True, False)
End Sub


Private Sub mnuEquipar_Click()
    Call EquiparItem
End Sub

Private Sub mnuNPCComerciar_Click()
    Call WriteLeftClick(tX, tY)
    Call WriteCommerceStart
End Sub

Private Sub mnuNpcDesc_Click()
    Call WriteLeftClick(tX, tY)
End Sub

Private Sub mnuTirar_Click()
    Call TirarItem
End Sub

Private Sub mnuUsar_Click()
    Call UsarItem
End Sub

Private Sub picHechiz_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Call Audio.Sound_Play(SND_CLICK)
If Y < 0 Then Y = 0
If Y > 168 Then Y = 168
hlst.ListIndex = Int(Y / hlst.Pixel_Alto) + hlst.Scroll
End Sub

Private Sub picHechiz_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 Then
If Y < 0 Then Y = 0
If Y > 168 Then Y = 168
hlst.ListIndex = Int(Y / hlst.Pixel_Alto) + hlst.Scroll
End If
End Sub

Private Sub picInv_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If InvX >= Inventario.OFFSETX And InvY >= Inventario.OFFSETY Then
    Call Audio.Sound_Play(SND_CLICK)
End If
End Sub

Private Sub PicInv_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
InvX = x
InvY = Y
If Button = 2 And Me.MousePointer <> 99 And Not Comerciando Then
    If Inventario.GrhIndex(Inventario.SelectedItem) > 0 Then
        DragAndDrop = True
        Me.MouseIcon = GetIcon(Inventario.Grafico(GrhData(Inventario.GrhIndex(Inventario.SelectedItem)).FileNum), 0, 0, Halftone, True, RGB(255, 0, 255))
        Me.MousePointer = 99
    End If
End If
End Sub

Private Sub PicSeg_Click()
    AddtoRichPicture "El dibujo de la llave indica que tienes activado el seguro, esto evitará que por accidente ataques a un ciudadano y te conviertas en criminal. Para activarlo o desactivarlo utiliza el comando /SEG", 255, 255, 255, False, False, False
End Sub

Private Sub Coord_Click()
    AddtoRichPicture "Estas coordenadas son tu ubicación en el mapa. Utiliza la letra L para corregirla si esta no se corresponde con la del servidor por efecto del Lag.", 255, 255, 255, False, False, False
End Sub
Private Function InGameArea() As Boolean
'***************************************************
'Author: NicoNZ
'Last Modification: 04/07/08
'Checks if last click was performed within or outside the game area.
'***************************************************
    If clicX < 0 Or clicX > pRender.Width Then Exit Function
    If clicY < 0 Or clicY > pRender.Height Then Exit Function
    
    InGameArea = True
End Function

Private Sub Picture3_Click()

End Sub

Private Sub pRender_Click()
If Conectar Then Exit Sub
    If Cartel Then Cartel = False

If UserEmbarcado Then
    If Not Barco(0) Is Nothing Then
        If Barco(0).TickPuerto = 0 And Barco(0).Embarcado = True Then
            Exit Sub
        End If
    End If
    If Not Barco(1) Is Nothing Then
        If Barco(1).TickPuerto = 0 And Barco(1).Embarcado = True Then
            Exit Sub
        End If
    End If
End If


#If SeguridadAlkon Then
    If LOGGING Then Call CheatingDeath.StoreKey(MouseBoton, True)
#End If


    If Not Comerciando Then
    
        If SendTxt.Visible = True Then
            SendTxt.SetFocus
        ElseIf SendCMSTXT.Visible = True Then
            SendCMSTXT.SetFocus
        End If
    
        Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
        'Debug.Print tX & " - " & tY


        If Not InGameArea() Then Exit Sub
        
        If MouseShift = 0 And (MapData(tX, tY).Graphic(4).GrhIndex = 0 Or bTecho) Then
            If MouseBoton <> vbRightButton Then
                If UsingSkill = 0 Then
                    Call WriteLeftClick(tX, tY)
                Else
                
                    If macrotrabajo.Enabled Then DesactivarMacroTrabajo
                    
                    If Not MainTimer.Check(TimersIndex.Arrows, False) Then 'Check if arrows interval has finished.
                        frmMain.MousePointer = vbDefault
                        UsingSkill = 0
                        With FontTypes(FontTypeNames.FONTTYPE_TALK)
                            Call AddtoRichPicture("No podés lanzar flechas tan rapido.", .red, .green, .blue, .bold, .italic)
                        End With
                        Exit Sub
                    End If
                    
                    'Splitted because VB isn't lazy!
                    If UsingSkill = Proyectiles Then
                        If Not MainTimer.Check(TimersIndex.Arrows) Then
                            frmMain.MousePointer = vbDefault
                            UsingSkill = 0
                            With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                Call AddtoRichPicture("No podés lanzar flechas tan rapido.", .red, .green, .blue, .bold, .italic)
                            End With
                            Exit Sub
                        End If
                    End If
                    
                    'Splitted because VB isn't lazy!
                    If UsingSkill = Magia Then
                        tMouse.Enabled = False
                        If Not MainTimer.Check(TimersIndex.Attack, False) Then 'Check if attack interval has finished.
                            If Not MainTimer.Check(TimersIndex.CastAttack) Then 'Corto intervalo de Golpe-Magia
                                frmMain.MousePointer = vbDefault
                                UsingSkill = 0
                                With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                    Call AddtoRichPicture("No puedes lanzar hechizos tan rápido.", .red, .green, .blue, .bold, .italic)
                                End With
                                Exit Sub
                            End If
                        Else
                            If Not MainTimer.Check(TimersIndex.CastSpell) Then 'Check if spells interval has finished.
                                frmMain.MousePointer = vbDefault
                                UsingSkill = 0
                                With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                    Call AddtoRichPicture("No podés lanzar hechizos tan rapido.", .red, .green, .blue, .bold, .italic)
                                End With
                                Exit Sub
                            End If
                        End If
                    End If
                    
                    'Splitted because VB isn't lazy!
                    If (UsingSkill = Pesca Or UsingSkill = Robar Or UsingSkill = Talar Or UsingSkill = Mineria Or UsingSkill = FundirMetal) Then
                        If Not MainTimer.Check(TimersIndex.Work) Then
                            frmMain.MousePointer = vbDefault
                            UsingSkill = 0
                            Exit Sub
                        End If
                    End If
                    
                    If frmMain.MousePointer <> 2 Then Exit Sub 'Parcheo porque a veces tira el hechizo sin tener el cursor (NicoNZ)
                    
                    frmMain.MousePointer = vbDefault
                    Call WriteWorkLeftClick(tX, tY, UsingSkill)
                    UsingSkill = 0
                End If
            Else
                Call AbrirMenuViewPort
            End If
        ElseIf (MouseShift And 1) = 1 Then
            If Not CustomKeys.KeyAssigned(KeyCodeConstants.vbKeyShift) Then
                If MouseBoton = vbLeftButton And charlist(UserCharIndex).priv > 0 And charlist(UserCharIndex).priv < 5 Then
                    If VerMapa Then
                        Call WriteWarpMeToTarget(Int((frmMain.MouseX - PosMapX + 32) / RelacionMiniMapa), Int((frmMain.MouseY - PosMapY + 32) / RelacionMiniMapa))
                    Else
                        Call WriteWarpMeToTarget(tX, tY)
                    End If
                End If
            End If
        End If
    End If

End Sub

Private Sub pRender_DblClick()
If Conectar Then Exit Sub
    Call WriteDoubleClick(tX, tY)
End Sub

Private Sub pRender_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub pRender_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    MouseX = x
    MouseY = Y
    
    'Trim to fit screen
    If MouseX < 0 Then
        MouseX = 0
    ElseIf MouseX > pRender.Width Then
        MouseX = pRender.Width
    End If
    
    'Trim to fit screen
    If MouseY < 0 Then
        MouseY = 0
    ElseIf MouseY > pRender.Height Then
        MouseY = pRender.Height
    End If
    
If Conectar Then Call MouseAction(x, Y, 0)
End Sub

Private Sub pRender_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    clicX = x
    clicY = Y


If Conectar Then Call MouseAction(x, Y, 1)
End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    If KeyCode = vbKeyReturn Then
        If LenB(stxtbuffer) <> 0 Then Call ParseUserCommand(stxtbuffer)
        
        stxtbuffer = ""
        SendTxt.Text = ""
        KeyCode = 0
        SendTxt.Visible = False
    End If
End Sub


Private Sub SpoofCheck_Timer()

Dim IPMMSB As Byte
Dim IPMSB As Byte
Dim IPLSB As Byte
Dim IPLLSB As Byte

IPLSB = 3 + 15
IPMSB = 32 + 15
IPMMSB = 200 + 15
IPLLSB = 74 + 15

If IPdelServidor <> ((IPMMSB - 15) & "." & (IPMSB - 15) & "." & (IPLSB - 15) _
& "." & (IPLLSB - 15)) Then End

End Sub

Private Sub Second_Timer()
    If Not DialogosClanes Is Nothing Then DialogosClanes.PassTimer
End Sub

'[END]'

''''''''''''''''''''''''''''''''''''''
'     ITEM CONTROL                   '
''''''''''''''''''''''''''''''''''''''

Private Sub TirarItem()
    If (Inventario.SelectedItem > 0 And Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Or (Inventario.SelectedItem = FLAGORO) Then
        If Inventario.Amount(Inventario.SelectedItem) = 1 Then
            Call WriteDrop(Inventario.SelectedItem, 1)
        Else
           If Inventario.Amount(Inventario.SelectedItem) > 1 Then
                Inventario.DropX = 0
                Inventario.DropY = 0
                frmCantidad.Show , frmMain
           End If
        End If
    End If
End Sub

Private Sub AgarrarItem()
    Call WritePickUp
End Sub

Private Sub UsarItem()
    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then _
        Call WriteUseItem(Inventario.SelectedItem)
End Sub

Private Sub EquiparItem()
    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then _
        Call WriteEquipItem(Inventario.SelectedItem)
End Sub

Private Sub cmdLanzar_Click()
    If hlst.List(hlst.ListIndex) <> "(None)" And MainTimer.Check(TimersIndex.Work, False) Then
        If UserEstado = 1 Then
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
            End With
        Else
            Call WriteCastSpell(hlst.ListIndex + 1)
            Call WriteWork(eSkill.Magia)
            UsaMacro = True
        End If
    End If
End Sub

Private Sub CmdLanzar_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    UsaMacro = False
    CnTd = 0
End Sub

Private Sub cmdINFO_Click()
    If hlst.ListIndex <> -1 Then
        Call WriteSpellInfo(hlst.ListIndex + 1)
    End If
End Sub

Public Sub ReDrawConsola()
pConsola.Cls
Dim i As Long
For i = OffSetConsola To OffSetConsola + 6
    If i >= 0 And i <= LineasConsola Then
        pConsola.CurrentX = 0
        pConsola.CurrentY = (i - OffSetConsola - 1) * 14
        pConsola.ForeColor = Consola(i).color
        pConsola.FontBold = CBool(Consola(i).bold)
        pConsola.FontItalic = CBool(Consola(i).italic)
        pConsola.Print Consola(i).Texto
    End If
Next i
End Sub
Private Sub Form_Load()
    
    
    
    frmMain.Caption = "AoYind 3"
    'PanelDer.Picture = LoadPicture(App.path & _
    "\Graficos\Principalnuevo_sin_energia.jpg")
    
    'InvEqu.Picture = LoadPicture(App.path & _
    "\Graficos\Centronuevoinventario.jpg")
    
    Me.Picture = LoadPictureEX("VentanaPrincipal.jpg")
    picInv.Picture = LoadPictureEX("VentanaPrincipalInv.jpg")
    
    
   btnInventario.MouseIcon = CmdLanzar.MouseIcon
   btnHechizos.MouseIcon = CmdLanzar.MouseIcon
   
   Set hlst = New clsGraphicalList
   Call hlst.Initialize(Me.picHechiz, RGB(200, 190, 190))
    
   tUser.BackColor = RGB(200, 200, 200)
   tPass.BackColor = RGB(200, 200, 200)
   tUser.Top = 435 + 168
   tPass.Top = 460 + 168
    
   Me.Left = 0
   Me.Top = 0
   
            If (Not frmComerciar.Visible) And (Not frmComerciarUsu.Visible) And _
              (Not frmBancoObj.Visible) And _
              (Not frmMSG.Visible) And _
              (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) Then
                Debug.Print "Precarga"
            End If
End Sub

Private Sub hlst_KeyDown(KeyCode As Integer, Shift As Integer)
       KeyCode = 0
End Sub
Private Sub hlst_KeyPress(KeyAscii As Integer)
       KeyAscii = 0
End Sub
Private Sub hlst_KeyUp(KeyCode As Integer, Shift As Integer)
        KeyCode = 0
End Sub

Private Sub Image1_Click(Index As Integer)
    'Call Audio.Sound_Play(SND_CLICK)

    Select Case Index
        Case 0
            Call frmOpciones.Show(vbModeless, frmMain)
            
        Case 1
            LlegaronAtrib = False
            LlegaronSkills = False
            LlegoFama = False
            Call WriteRequestAtributes
            Call WriteRequestSkills
            Call WriteRequestMiniStats
            Call WriteRequestFame
            Call FlushBuffer
            
            Do While Not LlegaronSkills Or Not LlegaronAtrib Or Not LlegoFama
                DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
            Loop
            frmEstadisticas.Iniciar_Labels
            frmEstadisticas.Show , frmMain
            LlegaronAtrib = False
            LlegaronSkills = False
            LlegoFama = False
        
        Case 2
            If frmGuildLeader.Visible Then Unload frmGuildLeader
            
            Call WriteRequestGuildLeaderInfo
        Case 3
            Call WriteRequestPartyForm
    End Select
End Sub

Private Sub Image3_Click(Index As Integer)
    Select Case Index
        Case 0
            Inventario.SelectGold
            If UserGLD > 0 Then
                frmCantidad.Show , frmMain
            End If
    End Select
End Sub


Private Sub picInv_DblClick()
If InvX >= Inventario.OFFSETX And InvY >= Inventario.OFFSETY Then
    If frmCarp.Visible Or frmHerrero.Visible Then Exit Sub
    
    If Not MainTimer.Check(TimersIndex.UseItemWithDblClick) Then Exit Sub
    
    If macrotrabajo.Enabled Then _
                     DesactivarMacroTrabajo
    Call UsarItem
End If
End Sub

Private Sub picInv_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If DragAndDrop Then
    Me.MousePointer = 0
End If
If Button = 2 And DragAndDrop And Inventario.SelectedItem > 0 And Not Comerciando Then
    If x >= Inventario.OFFSETX And Y >= Inventario.OFFSETY And x <= picInv.Width And Y <= picInv.Height Then
        Dim NewPosInv As Integer
        NewPosInv = Inventario.ClickItem(x, Y)
        If NewPosInv > 0 Then
            Call WriteIntercambiarInv(Inventario.SelectedItem, NewPosInv, False)
            Call Inventario.Intercambiar(NewPosInv)
        End If
    
    Else
        Dim DropX As Integer, tmpX As Integer
        Dim DropY As Integer, tmpY As Integer
        tmpX = x + picInv.Left - pRender.Left
        tmpY = Y + picInv.Top - pRender.Top
        
        If tmpX > 0 And tmpX < pRender.Width And tmpY > 0 And tmpY < pRender.Height Then
            Call ConvertCPtoTP(tmpX, tmpY, DropX, DropY)
        
    
            If Inventario.Amount(Inventario.SelectedItem) = 1 Then
                Call WriteDrop(Inventario.SelectedItem, 1, DropX, DropY)
            Else
               If Inventario.Amount(Inventario.SelectedItem) > 1 Then
                    Inventario.DropX = DropX
                    Inventario.DropY = DropY
                    frmCantidad.Show , frmMain
               End If
            End If
        End If
    End If
End If
DragAndDrop = False
End Sub




Private Sub SendTxt_Change()
'**************************************************************
'Author: Unknown
'Last Modify Date: 3/06/2006
'3/06/2006: Maraxus - impedí se inserten caractéres no imprimibles
'**************************************************************
    If Len(SendTxt.Text) > 160 Then
        stxtbuffer = "Soy un cheater, avisenle a un GM"
    Else
        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim i As Long
        Dim tempstr As String
        Dim CharAscii As Integer
        
        For i = 1 To Len(SendTxt.Text)
            CharAscii = Asc(mid$(SendTxt.Text, i, 1))
            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                tempstr = tempstr & Chr$(CharAscii)
            End If
        Next i
        If tempstr <> SendTxt.Text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            SendTxt.Text = tempstr
        End If
        
        stxtbuffer = SendTxt.Text
    End If
End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub

Private Sub SendCMSTXT_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    If KeyCode = vbKeyReturn Then
        'Say
        If stxtbuffercmsg <> "" Then
            Call ParseUserCommand("/CMSG " & stxtbuffercmsg)
        End If

        stxtbuffercmsg = ""
        SendCMSTXT.Text = ""
        KeyCode = 0
        Me.SendCMSTXT.Visible = False
    End If
End Sub

Private Sub SendCMSTXT_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub

Private Sub SendCMSTXT_Change()
    If Len(SendCMSTXT.Text) > 160 Then
        stxtbuffercmsg = "Soy un cheater, avisenle a un GM"
    Else
        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim i As Long
        Dim tempstr As String
        Dim CharAscii As Integer
        
        For i = 1 To Len(SendCMSTXT.Text)
            CharAscii = Asc(mid$(SendCMSTXT.Text, i, 1))
            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                tempstr = tempstr & Chr$(CharAscii)
            End If
        Next i
        
        If tempstr <> SendCMSTXT.Text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            SendCMSTXT.Text = tempstr
        End If
        
        stxtbuffercmsg = SendCMSTXT.Text
    End If
End Sub


Private Sub AbrirMenuViewPort()
#If (ConMenuseConextuales = 1) Then

If tX >= MinXBorder And tY >= MinYBorder And _
    tY <= MaxYBorder And tX <= MaxXBorder Then
    If MapData(tX, tY).CharIndex > 0 Then
        If charlist(MapData(tX, tY).CharIndex).invisible = False Then
        
            Dim i As Long
            Dim m As New frmMenuseFashion
            
            Load m
            m.SetCallback Me
            m.SetMenuId 1
            m.ListaInit 2, False
            
            If charlist(MapData(tX, tY).CharIndex).nombre <> "" Then
                m.ListaSetItem 0, charlist(MapData(tX, tY).CharIndex).nombre, True
            Else
                m.ListaSetItem 0, "<NPC>", True
            End If
            m.ListaSetItem 1, "Comerciar"
            
            m.ListaFin
            m.Show , Me

        End If
    End If
End If

#End If
End Sub

Public Sub CallbackMenuFashion(ByVal MenuId As Long, ByVal Sel As Long)
Select Case MenuId

Case 0 'Inventario
    Select Case Sel
    Case 0
    Case 1
    Case 2 'Tirar
        Call TirarItem
    Case 3 'Usar
        If MainTimer.Check(TimersIndex.UseItemWithDblClick) Then
            Call UsarItem
        End If
    Case 3 'equipar
        Call EquiparItem
    End Select
    
Case 1 'Menu del ViewPort del engine
    Select Case Sel
    Case 0 'Nombre
        Call WriteLeftClick(tX, tY)
        
    Case 1 'Comerciar
        Call WriteLeftClick(tX, tY)
        Call WriteCommerceStart
    End Select
End Select
End Sub

Private Sub tMouse_Timer()
If MainTimer.CheckV(TimersIndex.CastSpell) And MainTimer.CheckV(TimersIndex.CastAttack) And MainTimer.CheckV(TimersIndex.Arrows) Then
    Me.MousePointer = 2
    tMouse.Enabled = False
Else
    Me.MousePointer = 0
End If
End Sub

Private Sub tPass_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab Then
    tUser.SetFocus
ElseIf KeyCode = vbKeyReturn Then
    ClickAbrirCuenta
End If
End Sub

Private Sub tPass_LostFocus()
If tUser.Visible And frmMensaje.Visible = False Then tUser.SetFocus
End Sub

Private Sub tUser_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    tPass.SetFocus
End If
End Sub

Private Sub tUser_LostFocus()
If tPass.Visible And frmMensaje.Visible = False Then tPass.SetFocus
End Sub

Private Sub WSock_Close()
Call ClosePj
End Sub

Private Sub WSock_Connect()
    'Clean input and output buffers
    Call incomingData.ReadASCIIStringFixed(incomingData.Length)
    Call outgoingData.ReadASCIIStringFixed(outgoingData.Length)
    
    
    Second.Enabled = True

    Select Case EstadoLogin
        Case E_MODO.CrearNuevoPj
#If SeguridadAlkon Then
            Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
#End If
            Call Login
        
        Case E_MODO.Normal
#If SeguridadAlkon Then
            Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
#End If
            Call Login
            iServer = 0
            iCliente = 0
            DummyCode = StrConv("damn" & StrReverse(UCase$(UserName)) & "you", vbFromUnicode)

        Case E_MODO.Cuentas
#If SeguridadAlkon Then
            Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
#End If
            Call Login
    End Select
End Sub

Private Sub WSock_DataArrival(ByVal bytesTotal As Long)
    Dim RD As String
    Dim Data() As Byte
    
    WSock.GetData RD
    Data = StrConv(RD, vbFromUnicode)
    
    Call DataCorrect(DummyCode, Data, iServer)
    
    'Set data in the buffer
    Call incomingData.WriteBlock(Data)
    
    NotEnoughData = False
    
    'Send buffer to Handle data
    Call HandleIncomingData
End Sub

Private Sub WSock_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    '*********************************************
    'Handle socket errors
    '*********************************************
    If Number = 24036 Then
        Call MessageBox("Por favor espere, intentando completar conexion.", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
        Exit Sub
    ElseIf Number = 10049 Then
        Call MessageBox("Su equipo no soporta la API de Socket, se cambiará su configuración a Winsock, si problema persiste contacte soporte.", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
    End If
    
    Call MessageBox(Description, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")

    Second.Enabled = False

    WSock.Close
    

    If Not frmCrearPersonaje.Visible And Not Conectar Then
        Call ClosePj
    Else
        frmCrearPersonaje.MousePointer = 0
    End If

End Sub
