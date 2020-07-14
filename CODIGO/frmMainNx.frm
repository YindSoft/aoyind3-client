VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{6A967C3A-24E5-47BD-9298-705322683301}#1.0#0"; "cswskax6.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   ".::AoYind::."
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
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
   Picture         =   "frmMainN.frx":0CCA
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
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
      Left            =   240
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   2160
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
      Left            =   240
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   2160
      Visible         =   0   'False
      Width           =   8160
   End
   Begin VB.PictureBox pRender 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6240
      Left            =   195
      ScaleHeight     =   416
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   544
      TabIndex        =   8
      Top             =   2055
      Width           =   8160
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
         Left            =   4740
         MaxLength       =   160
         PasswordChar    =   "*"
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Chat"
         Top             =   4740
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
         Left            =   4740
         MaxLength       =   160
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Chat"
         Top             =   4245
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
      Left            =   8535
      Picture         =   "frmMainN.frx":46C12
      ScaleHeight     =   265
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   219
      TabIndex        =   18
      Top             =   2565
      Width           =   3285
      Begin VB.Label lblItem 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   600
         TabIndex        =   19
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
      Left            =   9105
      Picture         =   "frmMainN.frx":4DC27
      ScaleHeight     =   175
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   140
      TabIndex        =   17
      Top             =   3060
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
      Left            =   11205
      Picture         =   "frmMainN.frx":5FB85
      ScaleHeight     =   173
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   15
      Top             =   3090
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
         Picture         =   "frmMainN.frx":61C37
         ScaleHeight     =   7
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   14
         TabIndex        =   16
         Top             =   240
         Width           =   210
      End
   End
   Begin VB.Timer macrotrabajo 
      Enabled         =   0   'False
      Left            =   6480
      Top             =   840
   End
   Begin VB.PictureBox Picture1 
      Height          =   135
      Left            =   11715
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   6
      Top             =   8940
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Timer Macro 
      Interval        =   750
      Left            =   5520
      Top             =   840
   End
   Begin VB.Timer Second 
      Enabled         =   0   'False
      Interval        =   1050
      Left            =   5040
      Top             =   840
   End
   Begin VB.Timer SpoofCheck 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   4560
      Top             =   840
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   7560
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      RequestTimeout  =   30
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
      Left            =   255
      Picture         =   "frmMainN.frx":61DAD
      ScaleHeight     =   85
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   522
      TabIndex        =   4
      Top             =   480
      Width           =   7830
      Begin SocketWrenchCtl.SocketWrench SocketWrench1 
         Left            =   3360
         Top             =   840
         _cx             =   741
         _cy             =   741
      End
      Begin SocketWrenchCtl.SocketWrench Socket1 
         Left            =   6600
         Top             =   720
         _cx             =   741
         _cy             =   741
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
      Left            =   8145
      Picture         =   "frmMainN.frx":63DB1
      ScaleHeight     =   91
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   2
      Top             =   465
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
         Picture         =   "frmMainN.frx":651DB
         ScaleHeight     =   7
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   14
         TabIndex        =   3
         Top             =   1020
         Width           =   210
      End
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
      Left            =   8790
      TabIndex        =   29
      Top             =   8295
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
      Left            =   8790
      TabIndex        =   28
      Top             =   7950
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
      Left            =   8790
      TabIndex        =   27
      Top             =   7620
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
      Left            =   8790
      TabIndex        =   26
      Top             =   7305
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
      Left            =   8760
      TabIndex        =   25
      Top             =   8280
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
      Left            =   8760
      TabIndex        =   24
      Top             =   7935
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
      Left            =   8760
      TabIndex        =   23
      Top             =   7605
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
      Left            =   8760
      TabIndex        =   22
      Top             =   7290
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
      Left            =   8790
      TabIndex        =   21
      Top             =   6960
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
      Left            =   8760
      TabIndex        =   20
      Top             =   6945
      Width           =   1455
   End
   Begin VB.Image btnHechizos 
      Height          =   750
      Left            =   10215
      MousePointer    =   99  'Custom
      Picture         =   "frmMainN.frx":65351
      Top             =   1980
      Width           =   1530
   End
   Begin VB.Image btnInventario 
      Height          =   750
      Left            =   8685
      MousePointer    =   99  'Custom
      Picture         =   "frmMainN.frx":6A7AE
      Top             =   1980
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Image iBEXPE 
      Height          =   135
      Left            =   11160
      Picture         =   "frmMainN.frx":6FBE8
      Top             =   1485
      Width           =   60
   End
   Begin VB.Image PicMH 
      Height          =   375
      Left            =   2040
      Picture         =   "frmMainN.frx":6FC96
      Stretch         =   -1  'True
      Top             =   8520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Coord 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "(000,000)"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   5160
      TabIndex        =   14
      Top             =   8520
      Width           =   825
   End
   Begin VB.Image PicSeg 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   360
      Picture         =   "frmMainN.frx":70AA8
      Stretch         =   -1  'True
      Top             =   8520
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   11520
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
      Left            =   9600
      TabIndex        =   13
      Top             =   975
      Width           =   120
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   9960
      TabIndex        =   12
      Top             =   1005
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   375
      Index           =   0
      Left            =   11640
      MouseIcon       =   "frmMainN.frx":70F60
      MousePointer    =   99  'Custom
      Top             =   3060
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   375
      Index           =   1
      Left            =   11640
      MouseIcon       =   "frmMainN.frx":710B2
      MousePointer    =   99  'Custom
      Top             =   2640
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CmdLanzar 
      Height          =   525
      Left            =   8760
      MouseIcon       =   "frmMainN.frx":71204
      MousePointer    =   99  'Custom
      Top             =   5880
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.Image cmdInfo 
      Height          =   525
      Left            =   10680
      MouseIcon       =   "frmMainN.frx":71356
      MousePointer    =   99  'Custom
      Top             =   5880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label GldLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10000000"
      ForeColor       =   &H0000C0C0&
      Height          =   195
      Left            =   7200
      TabIndex        =   11
      Top             =   8505
      Width           =   840
   End
   Begin VB.Image Image3 
      Height          =   315
      Index           =   0
      Left            =   6795
      Top             =   8460
      Width           =   390
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   0
      Left            =   10440
      MouseIcon       =   "frmMainN.frx":714A8
      MousePointer    =   99  'Custom
      Top             =   8280
      Width           =   1410
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   1
      Left            =   10440
      MouseIcon       =   "frmMainN.frx":715FA
      MousePointer    =   99  'Custom
      Top             =   7920
      Width           =   1410
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   2
      Left            =   10440
      MouseIcon       =   "frmMainN.frx":7174C
      MousePointer    =   99  'Custom
      Top             =   7560
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
      TabIndex        =   10
      Top             =   2400
      Width           =   30
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "El Yind"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   9435
      TabIndex        =   9
      Top             =   735
      Width           =   2145
   End
   Begin VB.Image iBEXP 
      Height          =   135
      Left            =   9195
      Picture         =   "frmMainN.frx":7189E
      Top             =   1485
      Width           =   2025
   End
   Begin VB.Image STAShp 
      Height          =   180
      Left            =   8790
      Picture         =   "frmMainN.frx":7273A
      Top             =   6945
      Width           =   1395
   End
   Begin VB.Image MANShp 
      Height          =   180
      Left            =   8790
      Picture         =   "frmMainN.frx":7304E
      Top             =   7290
      Width           =   1395
   End
   Begin VB.Image Hpshp 
      Height          =   180
      Left            =   8790
      Picture         =   "frmMainN.frx":73962
      Top             =   7605
      Width           =   1395
   End
   Begin VB.Image COMIDAsp 
      Height          =   180
      Left            =   8790
      Picture         =   "frmMainN.frx":74276
      Top             =   7935
      Width           =   1395
   End
   Begin VB.Image AGUAsp 
      Height          =   180
      Left            =   8790
      Picture         =   "frmMainN.frx":74B8A
      Top             =   8280
      Width           =   1395
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
Public IsPlaying As Byte

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wmsg As Long, ByVal wparam As Long, lparam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Dim PuedeMacrear As Boolean
Dim OldYConsola As Integer
Public hlst As clsGraphicalList

Private Sub BarraConsola_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim TempY As Integer
    Dim TamCon As Integer
    TempY = y - 3
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

Private Sub BarraHechiz_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim TempY As Integer
    Dim TamCon As Integer
    TempY = y - 3
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

Private Sub BarritaConsola_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
    OldYConsola = y
End If
End Sub

Private Sub BarritaConsola_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
    Dim TempY As Integer
    TempY = BarritaConsola.Top + (y - OldYConsola)
    If TempY < 16 Then TempY = 16
    If TempY > 68 Then TempY = 68
    If LineasConsola <= 6 Then TempY = 68
    OffSetConsola = Int((TempY - 16) * (LineasConsola - 6) / 52)
    BarritaConsola.Top = TempY
    ReDrawConsola
End If
End Sub

Private Sub BarritaHechiz_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
    hlst.OldY = y
End If
End Sub

Private Sub BarritaHechiz_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
    Dim TempY As Integer
    Dim MaxItems As Integer
    MaxItems = Int(picHechiz.Height / hlst.Pixel_Alto)
    TempY = BarritaHechiz.Top + (y - hlst.OldY)
    If TempY < 16 Then TempY = 16
    If TempY > 150 Then TempY = 150
    If hlst.ListCount <= MaxItems Then TempY = 150
    hlst.Scroll = Int((TempY - 16) * (hlst.ListCount - MaxItems) / 134)
    BarritaHechiz.Top = TempY
End If
End Sub

Private Sub btnHechizos_Click()
    Call Audio.PlayWave(SND_CLICK)
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
    Call Audio.PlayWave(SND_CLICK)
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
Public Sub DibujarMH()
PicMH.Visible = True
End Sub

Public Sub DesDibujarMH()
PicMH.Visible = False
End Sub

Public Sub DibujarSeguro()
PicSeg.Visible = True
End Sub

Public Sub DesDibujarSeguro()
PicSeg.Visible = False
End Sub



Private Sub Command1_Click()
'IPdelServidor = "localhost"
'PuertoDelServidor = "7222"

'frmOldPersonaje.Show vbModal
'MostrarEntrar = GTCPres
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case CustomKeys.BindedKey(eKeyType.mKeyVerMapa)
        VerMapa = True
    Case vbKeyEscape
        If Conectar Then
            If GTCPres < 10000 Then
                GTCInicial = GTCInicial - (10000 - GTCPres)
                Call EscucharMp3(10)
            ElseIf MostrarEntrar > 0 Then
                If GTCPres - MostrarEntrar > 1000 Then
                    MostrarEntrar = -GTCPres
                    tUser.Visible = False
                    tPass.Visible = False
                    Call Audio.PlayWave(SND_CADENAS)
                End If
            Else
                prgRun = False
            End If
        End If
End Select
End Sub
Public Sub SetRender(Full As Boolean)
If Full Then
    pRender.Move 0, 0, 800, 600
Else
    pRender.Move 13, 137, 544, 416
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
                Case CustomKeys.BindedKey(eKeyType.mKeyVerMapa)
                    VerMapa = False
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
                        Call WriteWork(eSkill.Ocultarse)
                    End If
                                    
                Case CustomKeys.BindedKey(eKeyType.mKeyDropObject)
                    Call TirarItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyUseObject)
                    If macrotrabajo.Enabled Then DesactivarMacroTrabajo
                        
                    If MainTimer.Check(TimersIndex.UseItemWithU) Then
                        Call UsarItem
                    End If
                
                Case CustomKeys.BindedKey(eKeyType.mKeyRequestRefresh)
                    If MainTimer.Check(TimersIndex.SendRPU) Then
                        Call WriteRequestPositionUpdate
                        Beep
                    End If
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleSafeMode)
                    If frmMain.PicSeg.Visible Then
                        AddtoRichPicture "Escribe /SEG para quitar el seguro", 255, 255, 255, False, False, False
                    Else
                        Call WriteSafeToggle
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
              (Not frmBancoObj.Visible) And (Not frmSkills3.Visible) And _
              (Not frmMSG.Visible) And (Not frmForo.Visible) And _
              (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) Then
                SendCMSTXT.Visible = True
                SendCMSTXT.SetFocus
            End If
        
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
                AddtoRichPicture "No tan rápido..!", 255, 255, 255, False, False, False
            Else
                Call WriteMeditate
                PuedeMacrear = False
            End If
        
        Case CustomKeys.BindedKey(eKeyType.mKeyCastSpellMacro)
            If UserEstado = 1 Then
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                End With
                Exit Sub
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
              (Not frmBancoObj.Visible) And (Not frmSkills3.Visible) And _
              (Not frmMSG.Visible) And (Not frmForo.Visible) And _
              (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) Then
                SendTxt.Visible = True
                SendTxt.SetFocus
            End If
            
    End Select
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
ReleaseCapture
SendMessage Me.hwnd, &HA1, 2, 0&
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

Private Sub picHechiz_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Call Audio.PlayWave(SND_CLICK)
hlst.ListIndex = Int(y / hlst.Pixel_Alto) + hlst.Scroll
End Sub

Private Sub PicMH_Click()
    AddtoRichPicture "Auto lanzar hechizos. Utiliza esta habilidad para entrenar únicamente. Para activarlo/desactivarlo utiliza F7.", 255, 255, 255, False, False, False
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

#If SeguridadAlkon Then
    If LOGGING Then Call CheatingDeath.StoreKey(MouseBoton, True)
#End If

    If Not Comerciando Then
        Call ConvertCPtoTP(MouseX, MouseY, tX, tY)

        If Not InGameArea() Then Exit Sub
        
        If MouseShift = 0 Then
            If MouseBoton <> vbRightButton Then
                '[ybarra]
                If UsaMacro Then
                    CnTd = CnTd + 1
                    If CnTd = 3 Then
                        Call WriteUseSpellMacro
                        CnTd = 0
                    End If
                    UsaMacro = False
                End If
                '[/ybarra]
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
                If MouseBoton = vbLeftButton Then
                    Call WriteWarpChar("YO", UserMap, tX, tY)
                End If
            End If
        End If
    End If

End Sub

Private Sub pRender_DblClick()
If Conectar Then Exit Sub
    Call WriteDoubleClick(tX, tY)
End Sub

Private Sub pRender_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub pRender_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MouseX = x
    MouseY = y
    
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
    
If Conectar Then Call MouseAction(x, y, 0)
End Sub

Private Sub pRender_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    clicX = x
    clicY = y

If Conectar Then Call MouseAction(x, y, 1)
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

Private Sub CmdLanzar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
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
        pConsola.ForeColor = Consola(i).Color
        pConsola.FontBold = CBool(Consola(i).bold)
        pConsola.FontItalic = CBool(Consola(i).italic)
        pConsola.Print Consola(i).Texto
    End If
Next i
End Sub
Private Sub Form_Load()
    
    
    
    frmMain.Caption = "AoYind" & " V " & App.Major & "." & _
    App.Minor & "." & App.Revision
    'PanelDer.Picture = LoadPicture(App.path & _
    "\Graficos\Principalnuevo_sin_energia.jpg")
    
    'InvEqu.Picture = LoadPicture(App.path & _
    "\Graficos\Centronuevoinventario.jpg")
    
   btnInventario.MouseIcon = CmdLanzar.MouseIcon
   btnHechizos.MouseIcon = CmdLanzar.MouseIcon
   
   Set hlst = New clsGraphicalList
   Call hlst.Initialize(Me.picHechiz, RGB(200, 190, 190))
    
   tUser.BackColor = RGB(200, 200, 200)
   tPass.BackColor = RGB(200, 200, 200)
   tUser.Top = 435
   tPass.Top = 460
    
   Me.Left = 0
   Me.Top = 0
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
    'Call Audio.PlayWave(SND_CLICK)

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

Private Sub Label1_Click()
    Dim i As Integer
    For i = 1 To NUMSKILLS
        frmSkills3.Text1(i).Caption = UserSkills(i)
    Next i
    Alocados = SkillPoints
    frmSkills3.Puntos.Caption = "Puntos:" & SkillPoints
    frmSkills3.Show , frmMain
End Sub


Private Sub picInv_DblClick()
    If frmCarp.Visible Or frmHerrero.Visible Then Exit Sub
    
    If Not MainTimer.Check(TimersIndex.UseItemWithDblClick) Then Exit Sub
    
    If macrotrabajo.Enabled Then _
                     DesactivarMacroTrabajo
    
    Call UsarItem
End Sub

Private Sub picInv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Call Audio.PlayWave(SND_CLICK)
End Sub




Private Sub SendTxt_Change()
'**************************************************************
'Author: Unknown
'Last Modify Date: 3/06/2006
'3/06/2006: Maraxus - impedí se inserten caractéres no imprimibles
'**************************************************************
    If Len(SendTxt.Text) > 160 Then
        stxtbuffer = "Soy un cheater, avisenle a un gm"
    Else
        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim i As Long
        Dim TempStr As String
        Dim CharAscii As Integer
        
        For i = 1 To Len(SendTxt.Text)
            CharAscii = Asc(mid$(SendTxt.Text, i, 1))
            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                TempStr = TempStr & Chr$(CharAscii)
            End If
        Next i
        If TempStr <> SendTxt.Text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            SendTxt.Text = TempStr
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
        Dim TempStr As String
        Dim CharAscii As Integer
        
        For i = 1 To Len(SendCMSTXT.Text)
            CharAscii = Asc(mid$(SendCMSTXT.Text, i, 1))
            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                TempStr = TempStr & Chr$(CharAscii)
            End If
        Next i
        
        If TempStr <> SendCMSTXT.Text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            SendCMSTXT.Text = TempStr
        End If
        
        stxtbuffercmsg = SendCMSTXT.Text
    End If
End Sub


''''''''''''''''''''''''''''''''''''''
'     SOCKET1                        '
''''''''''''''''''''''''''''''''''''''
#If UsarWrench = 1 Then

Private Sub Socket1_OnConnect()
   
    'Clean input and output buffers
    Call incomingData.ReadASCIIStringFixed(incomingData.length)
    Call outgoingData.ReadASCIIStringFixed(outgoingData.length)
    
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
        
        Case E_MODO.Dados
#If SeguridadAlkon Then
            Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
#End If
            frmCrearPersonaje.Show vbModal

        Case E_MODO.Cuentas
#If SeguridadAlkon Then
            Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
#End If
            Call Login
    End Select
End Sub

Private Sub Socket1_OnDisconnect()
    Dim i As Long
    
    Second.Enabled = False
    Connected = False
    
    'Socket1.Cleanup
    
   
    On Local Error Resume Next
    For i = 0 To Forms.Count - 1
        If Forms(i).Name <> Me.Name And Forms(i).Name <> frmCrearPersonaje.Name And Forms(i).Name <> frmPasswd.Name Then
            Unload Forms(i)
        End If
    Next i
    On Local Error GoTo 0
    
    If Not frmPasswd.Visible And Not frmCrearPersonaje.Visible And Not Conectar Then
        ShowConnect
    End If
        
    pausa = False
    UserMeditar = False
    
#If SeguridadAlkon Then
    LOGGING = False
    LOGSTRING = False
    LastPressed = 0
    LastMouse = False
    LastAmount = 0
#End If

    UserClase = 0
    UserSexo = 0
    UserRaza = 0
    UserHogar = 0
    UserEmail = ""
    
    For i = 1 To NUMSKILLS
        UserSkills(i) = 0
    Next i

    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i
    
    macrotrabajo.Enabled = False

    SkillPoints = 0
    Alocados = 0
End Sub

Private Sub Socket1_OnError(ByVal ErrorCode As Variant, ByVal Description As Variant)
    '*********************************************
    'Handle socket errors
    '*********************************************
    If ErrorCode = 24036 Then
        Call MsgBox("Por favor espere, intentando completar conexion.", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
        Exit Sub
    End If
    
    Call MsgBox(Description, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")

    Second.Enabled = False

    frmMain.Socket1.Disconnect
    

    If Not frmCrearPersonaje.Visible And Not Conectar Then
        ShowConnect
    Else
        frmCrearPersonaje.MousePointer = 0
    End If
End Sub

Private Sub Socket1_OnRead()
    Dim RD As String
    Dim data() As Byte
    
    Call Socket1.Read(RD)
    data = StrConv(RD, vbFromUnicode)
    
    If RD = vbNullString Then Exit Sub
    
#If SeguridadAlkon Then
    Call DataReceived(data)
#End If
    
    'Put data in the buffer
    Call incomingData.WriteBlock(data)
    
    'Send buffer to Handle data
    Call HandleIncomingData
End Sub


#End If

Private Sub AbrirMenuViewPort()
#If (ConMenuseConextuales = 1) Then

If tX >= MinXBorder And tY >= MinYBorder And _
    tY <= MaxYBorder And tX <= MaxXBorder Then
    If MapData(tX, tY, UserMap).CharIndex > 0 Then
        If charlist(MapData(tX, tY, UserMap).CharIndex).invisible = False Then
        
            Dim i As Long
            Dim m As New frmMenuseFashion
            
            Load m
            m.SetCallback Me
            m.SetMenuId 1
            m.ListaInit 2, False
            
            If charlist(MapData(tX, tY, UserMap).CharIndex).Nombre <> "" Then
                m.ListaSetItem 0, charlist(MapData(tX, tY, UserMap).CharIndex).Nombre, True
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

Private Sub tPass_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab Then
    tUser.SetFocus
ElseIf KeyCode = vbKeyReturn Then
    ClickAbrirCuenta
End If
End Sub

Private Sub tPass_LostFocus()
If tUser.Visible Then tUser.SetFocus
End Sub

Private Sub tUser_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    tPass.SetFocus
End If
End Sub

Private Sub tUser_LostFocus()
If tPass.Visible Then tPass.SetFocus
End Sub
