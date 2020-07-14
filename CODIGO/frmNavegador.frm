VERSION 5.00
Begin VB.Form frmNavegador 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   ScaleHeight     =   580
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   7680
      TabIndex        =   1
      Top             =   8400
      Width           =   1335
   End
   Begin VB.PictureBox WB 
      Height          =   8415
      Left            =   0
      ScaleHeight     =   8355
      ScaleWidth      =   8940
      TabIndex        =   0
      Top             =   0
      Width           =   9000
   End
End
Attribute VB_Name = "frmNavegador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Enum eTipo
    Crear = 1
    Recuperar = 2
    Borrar = 3
End Enum
Public TIPO As eTipo
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
If TIPO = Crear Then
    'WB.Navigate ("http://www.aoyind.com/crearcuenta.php")
End If
End Sub
