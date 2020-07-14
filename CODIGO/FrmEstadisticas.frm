VERSION 5.00
Begin VB.Form frmEstadisticas 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Estadisticas"
   ClientHeight    =   6675
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   6975
   ClipControls    =   0   'False
   Icon            =   "FrmEstadisticas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   445
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image bAgregar 
      Height          =   135
      Index           =   19
      Left            =   6540
      Top             =   5805
      Width           =   150
   End
   Begin VB.Image bAgregar 
      Height          =   135
      Index           =   18
      Left            =   6540
      Top             =   5535
      Width           =   150
   End
   Begin VB.Image bAgregar 
      Height          =   135
      Index           =   17
      Left            =   6540
      Top             =   5265
      Width           =   150
   End
   Begin VB.Image bAgregar 
      Height          =   135
      Index           =   16
      Left            =   6540
      Top             =   4995
      Width           =   150
   End
   Begin VB.Image bAgregar 
      Height          =   135
      Index           =   15
      Left            =   6540
      Top             =   4725
      Width           =   150
   End
   Begin VB.Image bAgregar 
      Height          =   135
      Index           =   14
      Left            =   6540
      Top             =   4470
      Width           =   150
   End
   Begin VB.Image bAgregar 
      Height          =   135
      Index           =   13
      Left            =   6540
      Top             =   4230
      Width           =   150
   End
   Begin VB.Image bAgregar 
      Height          =   135
      Index           =   12
      Left            =   6540
      Top             =   3975
      Width           =   150
   End
   Begin VB.Image bAgregar 
      Height          =   135
      Index           =   11
      Left            =   6540
      Top             =   3720
      Width           =   150
   End
   Begin VB.Image bAgregar 
      Height          =   135
      Index           =   10
      Left            =   6540
      Top             =   3450
      Width           =   150
   End
   Begin VB.Image bAgregar 
      Height          =   135
      Index           =   9
      Left            =   6540
      Top             =   3195
      Width           =   150
   End
   Begin VB.Image bAgregar 
      Height          =   135
      Index           =   8
      Left            =   6540
      Top             =   2940
      Width           =   150
   End
   Begin VB.Image bAgregar 
      Height          =   135
      Index           =   7
      Left            =   6540
      Top             =   2655
      Width           =   150
   End
   Begin VB.Image bAgregar 
      Height          =   135
      Index           =   6
      Left            =   6540
      Top             =   2400
      Width           =   150
   End
   Begin VB.Image bAgregar 
      Height          =   135
      Index           =   5
      Left            =   6540
      Top             =   2145
      Width           =   150
   End
   Begin VB.Image bAgregar 
      Height          =   135
      Index           =   4
      Left            =   6540
      Top             =   1890
      Width           =   150
   End
   Begin VB.Image bAgregar 
      Height          =   135
      Index           =   3
      Left            =   6540
      Top             =   1650
      Width           =   150
   End
   Begin VB.Image bAgregar 
      Height          =   135
      Index           =   2
      Left            =   6540
      Top             =   1380
      Width           =   150
   End
   Begin VB.Image bAgregar 
      Height          =   135
      Index           =   1
      Left            =   6540
      Top             =   1125
      Width           =   150
   End
   Begin VB.Image bAgregar 
      Height          =   135
      Index           =   0
      Left            =   6540
      Top             =   885
      Width           =   150
   End
   Begin VB.Image bQuitar 
      Height          =   135
      Index           =   19
      Left            =   5220
      Top             =   5790
      Width           =   150
   End
   Begin VB.Image bQuitar 
      Height          =   135
      Index           =   18
      Left            =   5220
      Top             =   5520
      Width           =   150
   End
   Begin VB.Image bQuitar 
      Height          =   135
      Index           =   17
      Left            =   5220
      Top             =   5250
      Width           =   150
   End
   Begin VB.Image bQuitar 
      Height          =   135
      Index           =   16
      Left            =   5220
      Top             =   4980
      Width           =   150
   End
   Begin VB.Image bQuitar 
      Height          =   135
      Index           =   15
      Left            =   5220
      Top             =   4695
      Width           =   150
   End
   Begin VB.Image bQuitar 
      Height          =   135
      Index           =   14
      Left            =   5220
      Top             =   4455
      Width           =   150
   End
   Begin VB.Image bQuitar 
      Height          =   135
      Index           =   13
      Left            =   5220
      Top             =   4200
      Width           =   150
   End
   Begin VB.Image bQuitar 
      Height          =   135
      Index           =   12
      Left            =   5220
      Top             =   3945
      Width           =   150
   End
   Begin VB.Image bQuitar 
      Height          =   135
      Index           =   11
      Left            =   5220
      Top             =   3690
      Width           =   150
   End
   Begin VB.Image bQuitar 
      Height          =   135
      Index           =   10
      Left            =   5220
      Top             =   3435
      Width           =   150
   End
   Begin VB.Image bQuitar 
      Height          =   135
      Index           =   9
      Left            =   5220
      Top             =   3180
      Width           =   150
   End
   Begin VB.Image bQuitar 
      Height          =   135
      Index           =   8
      Left            =   5220
      Top             =   2925
      Width           =   150
   End
   Begin VB.Image bQuitar 
      Height          =   135
      Index           =   7
      Left            =   5220
      Top             =   2655
      Width           =   150
   End
   Begin VB.Image bQuitar 
      Height          =   135
      Index           =   6
      Left            =   5220
      Top             =   2385
      Width           =   150
   End
   Begin VB.Image bQuitar 
      Height          =   135
      Index           =   5
      Left            =   5220
      Top             =   2130
      Width           =   150
   End
   Begin VB.Image bQuitar 
      Height          =   135
      Index           =   4
      Left            =   5220
      Top             =   1860
      Width           =   150
   End
   Begin VB.Image bQuitar 
      Height          =   135
      Index           =   3
      Left            =   5220
      Top             =   1620
      Width           =   150
   End
   Begin VB.Image bQuitar 
      Height          =   135
      Index           =   2
      Left            =   5220
      Top             =   1365
      Width           =   150
   End
   Begin VB.Image bQuitar 
      Height          =   135
      Index           =   1
      Left            =   5220
      Top             =   1110
      Width           =   150
   End
   Begin VB.Image bQuitar 
      Height          =   135
      Index           =   0
      Left            =   5220
      Top             =   870
      Width           =   150
   End
   Begin VB.Label lblSkillPts 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A2BAC6&
      Height          =   255
      Left            =   5580
      TabIndex        =   37
      Top             =   585
      Width           =   615
   End
   Begin VB.Image imgCerrar 
      Height          =   360
      Left            =   105
      Tag             =   "1"
      Top             =   6240
      Width           =   6810
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   20
      Left            =   5415
      Top             =   5805
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   19
      Left            =   5415
      Top             =   5535
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   18
      Left            =   5415
      Top             =   5265
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   17
      Left            =   5415
      Top             =   4995
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   16
      Left            =   5415
      Top             =   4725
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   15
      Left            =   5415
      Top             =   4470
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   14
      Left            =   5415
      Top             =   4230
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   13
      Left            =   5415
      Top             =   3975
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   12
      Left            =   5415
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   11
      Left            =   5415
      Top             =   3450
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   10
      Left            =   5415
      Top             =   3195
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   9
      Left            =   5415
      Top             =   2940
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   8
      Left            =   5415
      Top             =   2670
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   7
      Left            =   5415
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   6
      Left            =   5415
      Top             =   2145
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   5
      Left            =   5415
      Top             =   1890
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   4
      Left            =   5415
      Top             =   1635
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   3
      Left            =   5415
      Top             =   1380
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   2
      Left            =   5415
      Top             =   1125
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   1
      Left            =   5415
      Top             =   885
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A2BAC6&
      Height          =   195
      Index           =   5
      Left            =   1725
      TabIndex        =   36
      Top             =   5835
      Width           =   555
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A2BAC6&
      Height          =   195
      Index           =   4
      Left            =   885
      TabIndex        =   35
      Top             =   5625
      Width           =   1785
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A2BAC6&
      Height          =   195
      Index           =   3
      Left            =   1905
      TabIndex        =   34
      Top             =   5415
      Width           =   825
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A2BAC6&
      Height          =   195
      Index           =   2
      Left            =   1830
      TabIndex        =   33
      Top             =   5190
      Width           =   825
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A2BAC6&
      Height          =   195
      Index           =   1
      Left            =   2025
      TabIndex        =   32
      Top             =   4980
      Width           =   825
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A2BAC6&
      Height          =   195
      Index           =   0
      Left            =   1965
      TabIndex        =   31
      Top             =   4755
      Width           =   825
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A2BAC6&
      Height          =   195
      Index           =   20
      Left            =   4185
      TabIndex        =   30
      Top             =   5790
      Width           =   270
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A2BAC6&
      Height          =   195
      Index           =   19
      Left            =   4665
      TabIndex        =   29
      Top             =   5520
      Width           =   270
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A2BAC6&
      Height          =   195
      Index           =   18
      Left            =   4800
      TabIndex        =   28
      Top             =   5235
      Width           =   270
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A2BAC6&
      Height          =   195
      Index           =   17
      Left            =   4485
      TabIndex        =   27
      Top             =   4995
      Width           =   270
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A2BAC6&
      Height          =   195
      Index           =   16
      Left            =   4005
      TabIndex        =   26
      Top             =   4755
      Width           =   270
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A2BAC6&
      Height          =   195
      Index           =   15
      Left            =   3915
      TabIndex        =   25
      Top             =   4485
      Width           =   270
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A2BAC6&
      Height          =   195
      Index           =   14
      Left            =   4110
      TabIndex        =   24
      Top             =   4230
      Width           =   270
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A2BAC6&
      Height          =   195
      Index           =   13
      Left            =   3825
      TabIndex        =   23
      Top             =   3960
      Width           =   270
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A2BAC6&
      Height          =   195
      Index           =   12
      Left            =   3660
      TabIndex        =   22
      Top             =   3675
      Width           =   270
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A2BAC6&
      Height          =   195
      Index           =   7
      Left            =   1005
      TabIndex        =   21
      Top             =   3885
      Width           =   270
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A2BAC6&
      Height          =   195
      Index           =   6
      Left            =   915
      TabIndex        =   20
      Top             =   3660
      Width           =   270
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A2BAC6&
      Height          =   195
      Index           =   5
      Left            =   960
      TabIndex        =   19
      Top             =   3420
      Width           =   270
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A2BAC6&
      Height          =   195
      Index           =   4
      Left            =   1065
      TabIndex        =   18
      Top             =   3150
      Width           =   270
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A2BAC6&
      Height          =   195
      Index           =   2
      Left            =   1110
      TabIndex        =   17
      Top             =   2910
      Width           =   270
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A2BAC6&
      Height          =   195
      Index           =   1
      Left            =   1095
      TabIndex        =   16
      Top             =   2670
      Width           =   270
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A2BAC6&
      Height          =   195
      Index           =   11
      Left            =   4800
      TabIndex        =   15
      Top             =   3420
      Width           =   270
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A2BAC6&
      Height          =   195
      Index           =   10
      Left            =   3930
      TabIndex        =   14
      Top             =   3180
      Width           =   270
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A2BAC6&
      Height          =   195
      Index           =   9
      Left            =   3645
      TabIndex        =   13
      Top             =   2940
      Width           =   270
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A2BAC6&
      Height          =   195
      Index           =   8
      Left            =   4260
      TabIndex        =   12
      Top             =   2655
      Width           =   270
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A2BAC6&
      Height          =   195
      Index           =   7
      Left            =   3945
      TabIndex        =   11
      Top             =   2370
      Width           =   270
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A2BAC6&
      Height          =   195
      Index           =   6
      Left            =   3900
      TabIndex        =   10
      Top             =   2145
      Width           =   270
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A2BAC6&
      Height          =   195
      Index           =   5
      Left            =   3795
      TabIndex        =   9
      Top             =   1875
      Width           =   270
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A2BAC6&
      Height          =   195
      Index           =   4
      Left            =   4710
      TabIndex        =   8
      Top             =   1605
      Width           =   270
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A2BAC6&
      Height          =   195
      Index           =   3
      Left            =   4725
      TabIndex        =   7
      Top             =   1380
      Width           =   270
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A2BAC6&
      Height          =   195
      Index           =   2
      Left            =   3630
      TabIndex        =   6
      Top             =   1125
      Width           =   270
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A2BAC6&
      Height          =   195
      Index           =   1
      Left            =   3675
      TabIndex        =   5
      Top             =   900
      Width           =   270
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A2BAC6&
      Height          =   195
      Index           =   5
      Left            =   1470
      TabIndex        =   4
      Top             =   1890
      Width           =   180
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A2BAC6&
      Height          =   195
      Index           =   4
      Left            =   1110
      TabIndex        =   3
      Top             =   1650
      Width           =   180
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A2BAC6&
      Height          =   195
      Index           =   3
      Left            =   1410
      TabIndex        =   2
      Top             =   1395
      Width           =   180
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A2BAC6&
      Height          =   195
      Index           =   2
      Left            =   1110
      TabIndex        =   1
      Top             =   1140
      Width           =   180
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A2BAC6&
      Height          =   195
      Index           =   1
      Left            =   1020
      TabIndex        =   0
      Top             =   885
      Width           =   180
   End
End
Attribute VB_Name = "frmEstadisticas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
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

Private cBotonCerrar As clsGraphicalButton
Public LastPressed As clsGraphicalButton

Private Const ANCHO_BARRA As Byte = 73 'pixeles
Private Const BAR_LEFT_POS As Integer = 361 'pixeles

Public Sub Iniciar_Labels()
'Iniciamos los labels con los valores de los atributos y los skills
Dim i As Integer


For i = 1 To NUMATRIBUTOS
    Atri(i).Caption = UserAtributos(i)
Next

For i = 1 To NUMSKILLS
    Skills(i).Caption = UserSkills(i)
    CalcularBarra (i)
Next

For i = 0 To bQuitar.UBound
    bQuitar(i).Tag = "0"
    bAgregar(i).Tag = "0"
Next i

lblSkillPts.Caption = SkillPoints

Label4(1).Caption = UserReputacion.AsesinoRep
Label4(2).Caption = UserReputacion.BandidoRep
'Label4(3).Caption = "Burgues: " & UserReputacion.BurguesRep
Label4(4).Caption = UserReputacion.LadronesRep
Label4(5).Caption = UserReputacion.NobleRep
Label4(6).Caption = UserReputacion.PlebeRep

If UserReputacion.Promedio < 0 Then
    Label4(7).ForeColor = vbRed
    Label4(7).Caption = "Criminal"
Else
    Label4(7).ForeColor = vbBlue
    Label4(7).Caption = "Ciudadano"
End If

With UserEstadisticas
    Label6(0).Caption = .CriminalesMatados
    Label6(1).Caption = .CiudadanosMatados
    Label6(2).Caption = .UsuariosMatados
    Label6(3).Caption = .NpcsMatados
    Label6(4).Caption = .Clase
    Label6(5).Caption = .PenaCarcel
End With

End Sub
Sub CalcularBarra(i As Integer)
Dim Ancho As Single
Dim TmpS As Single
    TmpS = UserSkillsMod(i)
    Ancho = TmpS * ANCHO_BARRA / 100   'IIf(PorcentajeSkills(i) = 0, ANCHO_BARRA, (100 - PorcentajeSkills(i)) / 100 * ANCHO_BARRA)
    shpSkillsBar(i).Width = ANCHO_BARRA - Ancho
    shpSkillsBar(i).Left = BAR_LEFT_POS + Ancho
End Sub
Private Sub bAgregar_Click(Index As Integer)
Dim i As Integer
Call Audio.Sound_Play(SND_CLICKNEW)
If SkillPoints > 0 And SPLibres > 0 Then
    i = Index + 1
    If UserSkillsMod(i) < 100 Then
        SPLibres = SPLibres - 1
        lblSkillPts.Caption = SPLibres
        UserSkillsMod(i) = UserSkillsMod(i) + 1
        Skills(i).Caption = UserSkillsMod(i)
        Skills(i).ForeColor = vbRed
        CalcularBarra (i)
    End If
End If
End Sub

Private Sub bAgregar_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If bAgregar(Index).Tag = "0" Then
    LimpiarBtnsDer (Index)
End If
End Sub

Private Sub bQuitar_Click(Index As Integer)
Dim i As Integer
Call Audio.Sound_Play(SND_CLICKNEW)
If SkillPoints > 0 And SPLibres < SkillPoints Then
    i = Index + 1
    If UserSkillsMod(i) > UserSkills(i) Then
        SPLibres = SPLibres + 1
        lblSkillPts.Caption = SPLibres
        UserSkillsMod(i) = UserSkillsMod(i) - 1
        Skills(i).Caption = UserSkillsMod(i)
        Skills(i).ForeColor = vbRed
        CalcularBarra (i)
        If UserSkillsMod(i) = UserSkills(i) Then
            Skills(i).ForeColor = &HA2BAC6
        End If
    End If
End If
End Sub

Private Sub bQuitar_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If bQuitar(Index).Tag = "0" Then
    LimpiarBtnsIzq (Index)
End If
End Sub
Sub LimpiarBtnsIzq(Index As Integer)
Dim i As Integer
For i = 0 To bQuitar.UBound
    If i <> Index Then
        If bQuitar(i).Tag = "1" Then
            bQuitar(i).Picture = Nothing
            bQuitar(i).Tag = "0"
        End If
    Else
        bQuitar(i).Picture = LoadPictureEX("FlechaHoverIzq.gif")
        bQuitar(i).Tag = "1"
    End If
Next i
End Sub
Sub LimpiarBtnsDer(Index As Integer)
Dim i As Integer
For i = 0 To bAgregar.UBound
    If i <> Index Then
        If bAgregar(i).Tag = "1" Then
            bAgregar(i).Picture = Nothing
            bAgregar(i).Tag = "0"
        End If
    Else
        bAgregar(i).Picture = LoadPictureEX("FlechaHoverDer.gif")
        bAgregar(i).Tag = "1"
    End If
Next i
End Sub

Private Sub Form_Load()
    Call SetTranslucent(Me.hwnd, NTRANS_GENERAL)
    Me.Picture = LoadPictureEX("VentanaEstadisticas.jpg")
    
    Call LoadButtons
End Sub

Private Sub LoadButtons()
    
    
    
    
    
    Set cBotonCerrar = New clsGraphicalButton
    Set LastPressed = New clsGraphicalButton
    
    Call cBotonCerrar.Initialize(imgCerrar, "BotonCerrarEstadisticas.jpg", _
                                    "BotonCerrarRolloverEstadisticas.jpg", _
                                    "BotonCerrarClickEstadisticas.jpg", Me)

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then MoverVentana (Me.hwnd)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
LimpiarBtnsIzq (-1)
LimpiarBtnsDer (-1)
    LastPressed.ToggleToNormal
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
End Sub

Private Sub imgCerrar_Click()
    Dim skillChanges(NUMSKILLS) As Byte
    Dim HayCambio As Boolean
    Dim i As Long
    HayCambio = False
    For i = 1 To NUMSKILLS
        skillChanges(i) = UserSkillsMod(i) - UserSkills(i)
        If UserSkillsMod(i) <> UserSkills(i) Then HayCambio = True
        UserSkills(i) = UserSkillsMod(i)
    Next i
    If HayCambio Then Call WriteModifySkills(skillChanges())
Unload Me
End Sub

Private Sub imgCerrar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If imgCerrar.Tag = "1" Then
        imgCerrar.Picture = LoadPictureEX("BotonCerrarApretadoEstadisticas.jpg")
        imgCerrar.Tag = "0"
    End If

End Sub

