VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBarco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const TIEMPO_EN_PUERTO As Long = 20
Private Const ACELERACION As Byte = 60
Private Const VELOCIDAD As Single = 7

Dim RutaBarco() As Position
Dim Paso As Byte
Dim Pos As Position
Dim Direccion As Byte
Dim UPasos As Byte
Dim GrhBarco As Grh
Dim GrhIBarco As Integer
Dim Sentido As Byte

Dim BOffSetX As Single
Dim BOffSetY As Single
Dim pEmbarcado As Boolean
Dim EnMovimiento As Boolean
Public TickPuerto As Long
Dim Puerto As Byte

Dim CapitanIndex As Integer
Dim Pasajeros(0 To 4) As Integer

Dim FactorBarco As Single

''''DEBUGING''''
Dim PasoCount As Integer
Dim PasoTick As Long
Dim InitTick As Long

Public Property Let Embarcado(ByVal Value As Boolean)
    If Value Then
        UserEmbarcado = True
        UserPos.x = Pos.x
        UserPos.y = Pos.y
    End If
    pEmbarcado = Value
End Property
Public Property Get Embarcado() As Boolean
    Embarcado = pEmbarcado
End Property

Public Property Get x() As Integer
    x = Pos.x
End Property
Public Property Get y() As Integer
    y = Pos.y
End Property

Private Function NextPuerto() As Integer
Dim P As Integer
If Sentido = 0 Then
    P = Puerto + 1
    If P > NUM_PUERTOS Then P = 1
Else
    P = Puerto - 1
    If P < 1 Then P = NUM_PUERTOS
End If
NextPuerto = P
End Function

Public Sub Init(ByVal Ruta As String, ByVal mPaso As Byte, ByVal x As Integer, ByVal y As Integer, ByVal TiempoPuerto As Long, ByVal mSentido As Byte, ByRef mPasajeros() As Integer)
Dim Rutas() As String
Rutas = Split(Ruta, ";")
Dim i As Integer
UPasos = UBound(Rutas)
ReDim RutaBarco(0 To UPasos) As Position
For i = 0 To UPasos
    RutaBarco(i).x = Val(ReadField(1, Rutas(i), 44))
    RutaBarco(i).y = Val(ReadField(2, Rutas(i), 44))
Next i
Paso = mPaso
Pos.x = x
Pos.y = y
Sentido = mSentido

For i = IIf(Sentido = 0, 1, NUM_PUERTOS) To IIf(Sentido = 0, NUM_PUERTOS, 1) Step IIf(Sentido = 0, 1, -1)
    If Paso > Puertos(i).Paso(Sentido) Then
        Puerto = i
    End If
Next i
Puerto = NextPuerto

If TiempoPuerto > 0 Then
    TickPuerto = (GetTickCount() And &H7FFFFFFF) - TiempoPuerto
    EnMovimiento = False
Else
    EnMovimiento = True
End If

'Capitan
CapitanIndex = 10000 - Sentido
charlist(CapitanIndex).Body = BodyData(428)
charlist(CapitanIndex).Head = HeadData(502)
charlist(CapitanIndex).Alpha = 255

Pasajeros(0) = CapitanIndex
For i = 1 To 4
    Pasajeros(i) = mPasajeros(i)
    If Pasajeros(i) = UserCharIndex And UserCharIndex > 0 Then
        Embarcado = True
    End If
Next i

If Pos.x > RutaBarco(Paso).x Then
    Direccion = west
    Call InitGrh(GrhBarco, 10384)
ElseIf Pos.x < RutaBarco(Paso).x Then
    Direccion = east
    Call InitGrh(GrhBarco, 10388)
ElseIf Pos.y > RutaBarco(Paso).y Then
    Direccion = north
    Call InitGrh(GrhBarco, 10396)
Else 'If Pos.Y < RutaBarco(Paso).Y Then '''REVISAR
    Direccion = south
    Call InitGrh(GrhBarco, 10392)
End If
Call CalcularFactorBarco

InitTick = (GetTickCount() And &H7FFFFFFF)
End Sub

Public Function Actualizar(ByVal x As Integer, ByVal y As Integer) As Boolean
Actualizar = Abs(x - Pos.x) > 20 Or Abs(y - Pos.y) > 20
End Function

Public Sub AgregarPasajero(ByVal Num As Byte, ByVal CharIndex As Integer)
Pasajeros(Num) = CharIndex
If CharIndex = UserCharIndex And UserCharIndex > 0 Then
    Embarcado = True
End If
End Sub
Public Sub QuitarPasajero(ByVal Num As Byte)
If Pasajeros(Num) = UserCharIndex Then
    Embarcado = False
    UserEmbarcado = False
End If
Pasajeros(Num) = 0
End Sub

Private Sub CalcularFactorBarco()
Dim NP As Byte
Dim Dist As Integer, Dist2 As Integer

NP = Puertos(Puerto).Paso(Sentido)
Dist = Abs(RutaBarco(NP).x - Pos.x) + Abs(RutaBarco(NP).y - Pos.y)

If Sentido = 0 Then
    If Puerto - 1 = 0 Then
        NP = Puertos(NUM_PUERTOS).Paso(Sentido)
    Else
        NP = Puertos(Puerto - 1).Paso(Sentido)
    End If
Else
    If Puerto + 1 > NUM_PUERTOS Then
        NP = Puertos(1).Paso(Sentido)
    Else
        NP = Puertos(Puerto + 1).Paso(Sentido)
    End If
End If

'Sistema de aceleracion, era muy brusco sino cuando el barco llegaba ;)
Dist2 = Abs(RutaBarco(NP).x - Pos.x) + Abs(RutaBarco(NP).y - Pos.y)
If Dist <= ACELERACION Then
    FactorBarco = (Dist / ACELERACION) * (VELOCIDAD - 0.5) + 0.5
ElseIf Dist2 <= ACELERACION Then
    FactorBarco = (Dist2 / ACELERACION) * (VELOCIDAD - 0.5) + 0.5
Else
    FactorBarco = 7
End If
End Sub

Public Sub Avanzar()
Dim NP As Integer
Select Case Direccion
    Case north
        Pos.y = Pos.y - 1
    Case south
        Pos.y = Pos.y + 1
    Case west
        Pos.x = Pos.x - 1
    Case east
        Pos.x = Pos.x + 1
End Select

If Embarcado Then
    UserPos.x = Pos.x
    UserPos.y = Pos.y
End If

Call CalcularFactorBarco


If Pos.x = RutaBarco(Paso).x And Pos.y = RutaBarco(Paso).y Then
    If EsPuerto(Paso) Then
        TickPuerto = (GetTickCount() And &H7FFFFFFF)
        NP = NextPuerto
        
        Call Audio.Sound_Play(215, Pos.x, Pos.y)
        If Embarcado Then
            Call Dialogos.CreateDialog("¡Hemos llegado al puerto de " & Puertos(Puerto).nombre & "! Pueden descender aquí o seguir rumbo a " & Puertos(NP).nombre, CapitanIndex, 255, 255, 255)
        Else
            Call Dialogos.CreateDialog("¡Todos a bordo! Zarparemos hacia el puerto de " & Puertos(NP).nombre & " en unos instantes.", CapitanIndex, 255, 255, 255)
        End If
        Puerto = NP
        
    End If
    Paso = NextPaso
    
    If Pos.x > RutaBarco(Paso).x Then
        Direccion = west
        Call InitGrh(GrhBarco, 10384)
    ElseIf Pos.x < RutaBarco(Paso).x Then
        Direccion = east
        Call InitGrh(GrhBarco, 10388)
    ElseIf Pos.y > RutaBarco(Paso).y Then
        Direccion = north
        Call InitGrh(GrhBarco, 10396)
    ElseIf Pos.y < RutaBarco(Paso).y Then
        Direccion = south
        Call InitGrh(GrhBarco, 10392)
    End If
End If
'If Embarcado And EnMovimiento And TickPuerto = 0 Then MoveScreen (Direccion)

'Debug.Print Pos.X & "  -  " & Pos.Y
End Sub
Private Function NextPaso() As Byte
If Paso + 1 > UPasos Then
    NextPaso = 0
Else
    NextPaso = Paso + 1
End If
End Function
Private Function EsPuerto(ByVal Paso As Byte) As Boolean
Dim i As Integer
For i = 1 To NUM_PUERTOS
    If Puertos(i).Paso(Sentido) = Paso Then
        EsPuerto = True
        Exit For
    End If
Next i
End Function

Public Sub Calcular()
Dim Factor As Double

'Debuging
'If PasoTick = 0 Then
'    Debug.Print Paso & " - " & Pos.X & ", " & Pos.Y & " - " & IIf(TickPuerto > 0, (GetTickCount() And &H7FFFFFFF) - TickPuerto, 0)
'    PasoTick = (GetTickCount() And &H7FFFFFFF)
'ElseIf (GetTickCount() And &H7FFFFFFF) - PasoTick >= 57000 Then
'    PasoTick = 0
'End If
'If Paso = 1 And (GetTickCount() And &H7FFFFFFF) - InitTick > 60000 Then
'    Debug.Print (GetTickCount() And &H7FFFFFFF) - InitTick
'    End
'End If

If Abs(Pos.x - UserPos.x) > 25 Or Abs(Pos.y - UserPos.y) > 25 Then
    Set Barco(Sentido) = Nothing
    Exit Sub
End If


If TickPuerto > 0 Then 'Ahora el barco no zarpa hasta que el servidor le diga.
    EnMovimiento = False
    'If (GetTickCount() And &H7FFFFFFF) - TickPuerto < TIEMPO_EN_PUERTO * 1000 Then
    '    EnMovimiento = False
    'Else
    '    TickPuerto = 0
    '    EnMovimiento = True
    'End If
Else
    EnMovimiento = True
End If

If EnMovimiento Then
    Factor = timerTicksPerFrame * FactorBarco * ScrollPixelsPerFrameX
    
    If Abs(BOffSetX) >= 32 Then
        BOffSetX = BOffSetX - Sgn(BOffSetX) * 32
        BOffSetY = 0
        Avanzar
    ElseIf Abs(BOffSetY) >= 32 Then
        BOffSetY = BOffSetY - Sgn(BOffSetY) * 32
        BOffSetX = 0
        Avanzar
    End If
    Select Case Direccion
            Case north
                BOffSetY = BOffSetY - Factor
            Case south
                BOffSetY = BOffSetY + Factor
            Case west
                BOffSetX = BOffSetX - Factor
            Case east
                BOffSetX = BOffSetX + Factor
    End Select
    If Embarcado Then
        BarcoOffSetX = BOffSetX
        BarcoOffSetY = BOffSetY
    End If
End If
End Sub

Public Sub Render(ByVal TileX As Integer, ByVal TileY As Integer, ByVal PixelOffSetX As Single, ByVal PixelOffSetY As Single)


Dim x As Integer
Dim y As Integer
Dim OffX As Integer, AddX As Integer
Dim OffY As Integer, AddY As Integer

x = (Pos.x - TileX + 11) * 32 + PixelOffSetX + 256 + BOffSetX
y = (Pos.y - TileY + 8) * 32 + PixelOffSetY + 256 + BOffSetY + 20

'Call DrawGrh(GrhBarco, X, Y, 1, 1, 0)
Select Case Direccion
    Case east
        Call DrawGrhIndex(5447, x, y, 1, D3DColorRGBA(IluRGB.R, IluRGB.G, IluRGB.B, 255))
    Case west
        Call DrawGrhIndex(5446, x, y, 1, D3DColorRGBA(IluRGB.R, IluRGB.G, IluRGB.B, 255))
    Case north
        Call DrawGrhIndex(5497, x, y, 1, D3DColorRGBA(IluRGB.R, IluRGB.G, IluRGB.B, 255))
    Case south
        Call DrawGrhIndex(5496, x, y, 1, D3DColorRGBA(IluRGB.R, IluRGB.G, IluRGB.B, 255))
End Select

OffX = x
OffY = y
'Dibujamos a los pasajeros y capitan

Select Case Direccion
    Case north
        OffY = OffY - 155
        AddY = 20
    Case south
        OffY = OffY - 45
        AddY = -20
    Case east
        OffX = OffX + 67
        OffY = OffY - 44
        AddX = -20
    Case west
        OffX = OffX - 67
        OffY = OffY - 44
        AddX = 20
End Select

Dim i As Integer
Dim Index As Integer

For i = IIf(Direccion = north, 0, 4) To IIf(Direccion = north, 4, 0) Step IIf(Direccion = north, 1, -1)
    Index = Pasajeros(i)
    If Index > 0 Then
        charlist(Index).Heading = Direccion
        charlist(Index).priv = 10
        Call CharRender(charlist(Index), Index, OffX + AddX * IIf(i > 0 And (Direccion = east Or Direccion = west), 9.5 - i, 0), OffY + AddY * IIf(i > 0 And (Direccion = north Or Direccion = south), 6 - i, 0))
    End If
Next i

Select Case Direccion
    Case east
        Call DrawGrhIndex(5449, x - 108, y - 29, 0, D3DColorRGBA(IluRGB.R, IluRGB.G, IluRGB.B, 255))
    Case west
        Call DrawGrhIndex(5448, x + 56, y - 29, 0, D3DColorRGBA(IluRGB.R, IluRGB.G, IluRGB.B, 255))
    Case north
        Call DrawGrhIndex(5499, x, y, 1, D3DColorRGBA(IluRGB.R, IluRGB.G, IluRGB.B, 220))
    Case south
        Call DrawGrhIndex(5498, x + 2, y, 1, D3DColorRGBA(IluRGB.R, IluRGB.G, IluRGB.B, 220))
        Index = Pasajeros(0)
        Call CharRender(charlist(Index), Index, OffX, OffY)

End Select


End Sub
