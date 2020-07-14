VERSION 5.00
Begin VB.Form frmRetos 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4755
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   344
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   317
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox tOro 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   225
      Left            =   960
      TabIndex        =   4
      Text            =   "0"
      Top             =   1260
      Width           =   1455
   End
   Begin VB.TextBox j4 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A2BAC6&
      Height          =   225
      Left            =   2730
      TabIndex        =   3
      Text            =   "El Yind"
      Top             =   3600
      Width           =   1350
   End
   Begin VB.TextBox j3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A2BAC6&
      Height          =   225
      Left            =   645
      TabIndex        =   2
      Text            =   "El Yind"
      Top             =   3600
      Width           =   1350
   End
   Begin VB.TextBox j2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A2BAC6&
      Height          =   225
      Left            =   2730
      TabIndex        =   1
      Text            =   "El Yind"
      Top             =   2895
      Width           =   1350
   End
   Begin VB.TextBox j1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A2BAC6&
      Height          =   225
      Left            =   645
      TabIndex        =   0
      Text            =   "El Yind"
      Top             =   2895
      Width           =   1350
   End
   Begin VB.Image Aj4 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   3660
      Top             =   3195
      Width           =   300
   End
   Begin VB.Image Aj2 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   3660
      Top             =   2505
      Width           =   300
   End
   Begin VB.Image Aj3 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   1515
      Top             =   3195
      Width           =   300
   End
   Begin VB.Image Aj1 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   1515
      Top             =   2505
      Width           =   300
   End
   Begin VB.Image imgRechazar 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   1875
      Top             =   4410
      Width           =   1455
   End
   Begin VB.Image imgSalir 
      Height          =   315
      Left            =   3480
      Top             =   4425
      Width           =   930
   End
   Begin VB.Image imgAceptar 
      Height          =   360
      Left            =   345
      Top             =   4410
      Width           =   1455
   End
   Begin VB.Image chkItems 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   2670
      Top             =   1170
      Width           =   300
   End
   Begin VB.Image opt2vs2 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   2670
      Top             =   735
      Width           =   300
   End
   Begin VB.Image opt1vs1 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   1875
      Top             =   735
      Width           =   300
   End
End
Attribute VB_Name = "frmRetos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Vs As Byte
Dim PorItems As Boolean
Dim Oro As Long
Dim CreaReto As Boolean
Private cBotonAceptar As clsGraphicalButton
Private cBotonRechazar As clsGraphicalButton
Private cBotonSalir As clsGraphicalButton

Public LastPressed As clsGraphicalButton
Private Sub chkItems_Click()
Call Audio.Sound_Play(SND_CLICKNEW)
PorItems = Not PorItems
If PorItems Then
    chkItems.Picture = LoadPictureEX("Tilde.gif")
Else
    chkItems.Picture = Nothing
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    imgSalir_Click
End If
End Sub

Private Sub Form_Load()
Call SetTranslucent(Me.hwnd, NTRANS_GENERAL)
    
    Me.Picture = LoadPictureEX("VentanaRetos.jpg")
    
    Call LoadButtons
End Sub
Public Sub Iniciar(pVs As Byte, pPorItems As Boolean, pOro As Long, Pj1 As String, Apj1 As Boolean, Pj2 As String, Apj2 As Boolean, Pj3 As String, Apj3 As Boolean, Pj4 As String, Apj4 As Boolean, Acepto As Boolean)
Vs = pVs
PorItems = pPorItems
Oro = pOro
j1.Text = Pj1
j2.Text = Pj2
j3.Text = Pj3
j4.Text = Pj4
tOro.Text = Oro
If Vs = 1 Then
    opt1vs1.Picture = LoadPictureEX("Tilde.gif")
Else
    opt2vs2.Picture = LoadPictureEX("Tilde.gif")
End If

If PorItems Then chkItems.Picture = LoadPictureEX("Tilde.gif")

If Apj1 Then Aj1.Picture = LoadPictureEX("Tilde.gif")
If Apj2 Then Aj2.Picture = LoadPictureEX("Tilde.gif")
If Apj3 Then Aj3.Picture = LoadPictureEX("Tilde.gif")
If Apj4 Then Aj4.Picture = LoadPictureEX("Tilde.gif")

CreaReto = Pj2 = ""

tOro.Locked = Not CreaReto
j1.Locked = True
j2.Locked = Not CreaReto
j3.Locked = Not CreaReto
j4.Locked = Not CreaReto
j3.Enabled = Vs = 2
j4.Enabled = Vs = 2
imgRechazar.Enabled = Not CreaReto
cBotonRechazar.EnableButton imgRechazar.Enabled
imgAceptar.Enabled = Not Acepto
cBotonAceptar.EnableButton imgAceptar.Enabled
opt1vs1.Enabled = CreaReto
opt2vs2.Enabled = CreaReto
chkItems.Enabled = CreaReto
End Sub

Private Sub imgAceptar_Click()
If CreaReto Then
    'Prevalidaciones del cliente
    If Trim$(j2.Text) = "" Then
        frmMensaje.msg.Caption = "Debe ingresar el nombre del jugador 2."
        frmMensaje.Show
        Exit Sub
    End If
    If Vs = 2 Then
        If Trim$(j3.Text) = "" Or Trim$(j4.Text) = "" Then
            frmMensaje.msg.Caption = "Debe ingresar el nombre del jugador 3 y 4."
            frmMensaje.Show
            Exit Sub
        End If
    End If
    If Oro <= 0 Then
        frmMensaje.msg.Caption = "Ingrese el monto de la apuesta."
        frmMensaje.Show
        Exit Sub
    End If
    Call WriteRetosCrear(Vs, PorItems, Oro, j2.Text, j3.Text, j4.Text)
Else
    Call WriteRetosDecide(True)
End If
End Sub

Private Sub imgRechazar_Click()
Call WriteRetosDecide(False)
End Sub

Private Sub imgSalir_Click()
Unload Me
End Sub

Private Sub opt1vs1_Click()
Call Audio.Sound_Play(SND_CLICKNEW)

If Vs <> 1 Then
    opt1vs1.Picture = LoadPictureEX("Tilde.gif")
    opt2vs2.Picture = Nothing
    Vs = 1
    j3.Text = ""
    j3.Enabled = False
    j4.Text = ""
    j4.Enabled = False
End If
End Sub

Private Sub opt2vs2_Click()
Call Audio.Sound_Play(SND_CLICKNEW)

If Vs <> 2 Then
    opt2vs2.Picture = LoadPictureEX("Tilde.gif")
    opt1vs1.Picture = Nothing
    Vs = 2
    j3.Enabled = True
    j4.Enabled = True
End If
End Sub
Private Sub LoadButtons()

    Set cBotonAceptar = New clsGraphicalButton
    Set cBotonRechazar = New clsGraphicalButton
    Set cBotonSalir = New clsGraphicalButton
    
    Set LastPressed = New clsGraphicalButton
    
    
    Call cBotonAceptar.Initialize(imgAceptar, "BotonAceptarComUsu.jpg", _
                                        "BotonAceptarRolloverComUsu.jpg", _
                                        "BotonAceptarClickComUsu.jpg", Me, _
                                        "BotonAceptarGrisComUsu.jpg", True)
                                        
    Call cBotonRechazar.Initialize(imgRechazar, "BotonRechazarComUsu.jpg", _
                                        "BotonRechazarRolloverComUsu.jpg", _
                                        "BotonRechazarClickComUsu.jpg", Me, _
                                        "BotonRechazarGrisComUsu.jpg", True)
    
    Call cBotonSalir.Initialize(imgSalir, "BotonSalirAlineacion.jpg", _
                                    "BotonSalirRolloverAlineacion.jpg", _
                                    "BotonSalirClickAlineacion.jpg", Me)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then MoverVentana (Me.hwnd)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    LastPressed.ToggleToNormal
End Sub

Private Sub tOro_Change()
Oro = Val(tOro.Text)
End Sub
