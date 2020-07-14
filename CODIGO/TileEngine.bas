Attribute VB_Name = "Mod_TileEngine"
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
Public Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hDCDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hDCSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal crTransparent As Long) As Long

'Map sizes in tiles
Public Const XMaxMapSize As Integer = 1100
Public Const XMinMapSize As Integer = 1
Public Const YMaxMapSize As Integer = 1500
Public Const YMinMapSize As Integer = 1

Public Const RelacionMiniMapa As Single = 1.92120075046904

Public Const GrhFogata As Integer = 1521

''
'Sets a Grh animation to loop indefinitely.
Private Const INFINITE_LOOPS As Integer = -1


'Encabezado bmp
Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type

'Info del encabezado del bmp
Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

'Posicion en un mapa
Public Type Position
    x As Long
    Y As Long
End Type

'Posicion en el Mundo
Public Type WorldPos
    Map As Integer
    x As Integer
    Y As Integer
End Type

'Contiene info acerca de donde se puede encontrar un grh tamaño y animacion
Type GrhData
    sX As Integer
    sY As Integer
    
    FileNum As Integer
    
    pixelWidth As Integer
    pixelHeight As Integer
    
    TileWidth As Single
    TileHeight As Single
    
    NumFrames As Integer
    Frames() As Integer
    
    Speed As Single
End Type

'apunta a una estructura grhdata y mantiene la animacion
Type Grh
    GrhIndex As Integer
    FrameCounter As Single
    Speed As Single
    Started As Byte
    Loops As Integer
End Type

'Lista de cuerpos
Type BodyData
    Walk(E_Heading.north To E_Heading.west) As Grh
    HeadOffset As Position
End Type

'Lista de cabezas
Type HeadData
    Head(E_Heading.north To E_Heading.west) As Grh
End Type

'Lista de las animaciones de las armas
Type WeaponAnimData
    WeaponWalk(E_Heading.north To E_Heading.west) As Grh
    '[ANIM ATAK]
    WeaponAttack As Byte
End Type

'Lista de las animaciones de los escudos
Type ShieldAnimData
    ShieldWalk(E_Heading.north To E_Heading.west) As Grh
    '[ANIM ATAK]
    ShieldAttack As Byte
End Type

Public NPCMuertos As New Collection

'Apariencia del personaje
Public Type Char
    ACTIVE As Byte
    Heading As E_Heading
    Pos As Position
    
    iHead As Integer
    iBody As Integer
    Body As BodyData
    Head As HeadData
    Casco As HeadData
    Arma As WeaponAnimData
    Escudo As ShieldAnimData
    UsandoArma As Boolean
    
    fX As Grh
    FxIndex As Integer
    
    Criminal As Byte
    
    nombre As String
    
    scrollDirectionX As Integer
    scrollDirectionY As Integer
    
    Moving As Byte
    MoveOffsetX As Single
    MoveOffsetY As Single
    
    pie As Boolean
    logged As Boolean
    muerto As Boolean
    invisible As Boolean
    Alpha As Byte
    ContadorInvi As Integer
    iTick As Long
    priv As Byte
    
    Quieto As Byte
End Type

'Info de un objeto
Public Type Obj
    OBJIndex As Integer
    Amount As Integer
End Type

'Tipo de las celdas del mapa
Public Type MapBlock
    Graphic(1 To 4) As Grh
    CharIndex As Integer
    ObjGrh As Grh
    
    Particle_Group_Index As Integer
    
    Blocked As Byte
    Trigger As Byte
    
    Map As Byte
    Elemento As Object
    
    Light_Value(3) As Long
    Hora As Byte
End Type

Public IniPath As String
Public MapPath As String


'Bordes del mapa
Public MinXBorder As Integer
Public MaxXBorder As Integer
Public MinYBorder As Integer
Public MaxYBorder As Integer

'Status del user
Public CurMap As Integer 'Mapa actual
Public UserIndex As Integer
Public UserMoving As Byte
Public UserBody As Integer
Public UserHead As Integer
Public UserPos As Position 'Posicion
Public AddtoUserPos As Position 'Si se mueve
Public UserCharIndex As Integer

Public EngineRun As Boolean

Public FPS As Long
Public FramesPerSecCounter As Long
Private fpsLastCheck As Long

'Tamaño del la vista en Tiles
Private WindowTileWidth As Integer
Private WindowTileHeight As Integer

Private HalfWindowTileWidth As Integer
Private HalfWindowTileHeight As Integer

'Offset del desde 0,0 del main view
Private MainViewTop As Integer
Private MainViewLeft As Integer

'Cuantos tiles el engine mete en el BUFFER cuando
'dibuja el mapa. Ojo un tamaño muy grande puede
'volver el engine muy lento
Public TileBufferSize As Integer

Private TileBufferPixelOffsetX As Integer
Private TileBufferPixelOffsetY As Integer

'Tamaño de los tiles en pixels
Public TilePixelHeight As Integer
Public TilePixelWidth As Integer

'Number of pixels the engine scrolls per frame. MUST divide evenly into pixels per tile
Public ScrollPixelsPerFrameX As Integer
Public ScrollPixelsPerFrameY As Integer

Public timerElapsedTime As Single
Public timerTicksPerFrame As Single
Public engineBaseSpeed As Single


Public NumBodies As Integer
Public Numheads As Integer
Public NumFxs As Integer
Public NumWeaponAnims As Integer
Public NumShieldAnims As Integer

Public NumChars As Integer
Public LastChar As Integer


Private MainDestRect   As RECT
Private MainViewRect   As RECT
Private BackBufferRect As RECT

Private MainViewWidth As Integer
Private MainViewHeight As Integer

Private MouseTileX As Integer
Private MouseTileY As Integer




'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Graficos¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public GrhData() As GrhData 'Guarda todos los grh
Public BodyData() As BodyData
Public HeadData() As HeadData
Public FxData() As tIndiceFx
Public WeaponAnimData() As WeaponAnimData
Public ShieldAnimData() As ShieldAnimData
Public CascoAnimData() As HeadData
Public Arrojas As New Collection
Public Tooltips As New Collection
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Mapa?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock ' Mapa
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

Public bRain        As Boolean 'está raineando?
Public bTecho       As Boolean 'hay techo?
Public bAlpha       As Byte
Public tTick        As Long
Public ColorTecho   As Long
Public brstTick     As Long

Private RLluvia(7)  As RECT  'RECT de la lluvia
Private iFrameIndex As Byte  'Frame actual de la LL
Private llTick      As Long  'Contador
Private LTLluvia(4) As Integer

Public charlist(1 To 10000) As Char
Public AperturaPergamino As Single

#If SeguridadAlkon Then

Public MI(1 To 1233) As clsManagerInvisibles
Public CualMI As Integer

#End If

' Used by GetTextExtentPoint32
Private Type Size
    cx As Long
    cy As Long
End Type

'[CODE 001]:MatuX
Public Enum PlayLoop
    plNone = 0
    plLluviain = 1
    plLluviaout = 2
End Enum
'[END]'
'
'       [END]
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public MapaY As Single
Public VerMapa As Boolean
Public Entrada As Byte
Public FrameUseMotionBlur As Boolean

'Very percise counter 64bit system counter
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'Text width computation. Needed to center text.
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long

Public PosMapX As Single
Public PosMapY As Single

Sub CargarCabezas()
    Dim N As Integer
    Dim i As Long
    Dim Numheads As Integer
    Dim Miscabezas() As tIndiceCabeza
    
    N = FreeFile()
    Open App.path & "\init\Cabezas.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , Numheads
    
    'Resize array
    ReDim HeadData(0 To Numheads) As HeadData
    ReDim Miscabezas(0 To Numheads) As tIndiceCabeza
    
    For i = 1 To Numheads
        Get #N, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call InitGrh(HeadData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call InitGrh(HeadData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call InitGrh(HeadData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call InitGrh(HeadData(i).Head(4), Miscabezas(i).Head(4), 0)
        End If
    Next i
    
    Close #N
End Sub

Sub CargarCascos()
    Dim N As Integer
    Dim i As Long
    Dim NumCascos As Integer

    Dim Miscabezas() As tIndiceCabeza
    
    N = FreeFile()
    Open App.path & "\init\Cascos.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCascos
    
    'Resize array
    ReDim CascoAnimData(0 To NumCascos) As HeadData
    ReDim Miscabezas(0 To NumCascos) As tIndiceCabeza
    
    For i = 1 To NumCascos
        Get #N, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call InitGrh(CascoAnimData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call InitGrh(CascoAnimData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call InitGrh(CascoAnimData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call InitGrh(CascoAnimData(i).Head(4), Miscabezas(i).Head(4), 0)
        End If
    Next i
    
    Close #N
End Sub

Sub CargarCuerpos()
    Dim N As Integer
    Dim i As Long
    Dim NumCuerpos As Integer
    Dim MisCuerpos() As tIndiceCuerpo
    
    N = FreeFile()
    Open App.path & "\init\Personajes.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCuerpos
    
    'Resize array
    ReDim BodyData(0 To NumCuerpos) As BodyData
    ReDim MisCuerpos(0 To NumCuerpos) As tIndiceCuerpo
    
    For i = 1 To NumCuerpos
        Get #N, , MisCuerpos(i)
        
        If MisCuerpos(i).Body(1) Then
            InitGrh BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0
            InitGrh BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0
            InitGrh BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0
            InitGrh BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0
            
            BodyData(i).HeadOffset.x = MisCuerpos(i).HeadOffsetX
            BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY
        End If
    Next i
    
    Close #N
End Sub

Sub CargarFxs()
    Dim N As Integer
    Dim i As Long
    Dim NumFxs As Integer
    
    N = FreeFile()
    Open App.path & "\init\Fxs.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumFxs
    
    'Resize array
    ReDim FxData(1 To NumFxs) As tIndiceFx
    
    For i = 1 To NumFxs
        Get #N, , FxData(i)
    Next i
    
    Close #N
End Sub


Sub ConvertCPtoTP(ByVal viewPortX As Integer, ByVal viewPortY As Integer, ByRef tX As Integer, ByRef tY As Integer)
'******************************************
'Converts where the mouse is in the main window to a tile position. MUST be called eveytime the mouse moves.
'******************************************
    tX = UserPos.x + viewPortX \ TilePixelWidth - WindowTileWidth \ 2
    tY = UserPos.Y + viewPortY \ TilePixelHeight - WindowTileHeight \ 2
End Sub

Sub MakeChar(ByVal CharIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal x As Integer, ByVal Y As Integer, ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer)
On Error Resume Next
    'Apuntamos al ultimo Char
    If CharIndex > LastChar Then LastChar = CharIndex
    
    With charlist(CharIndex)
        'If the char wasn't allready active (we are rewritting it) don't increase char count
        If .ACTIVE = 0 Then _
            NumChars = NumChars + 1
        
        If Arma = 0 Then Arma = 2
        If Escudo = 0 Then Escudo = 2
        If Casco = 0 Then Casco = 2
        
        .iHead = Head
        .iBody = Body
        .Head = HeadData(Head)
        .Body = BodyData(Body)
        .Arma = WeaponAnimData(Arma)
        '[ANIM ATAK]
        .Arma.WeaponAttack = 0
        .Escudo.ShieldAttack = 0
        
        .Escudo = ShieldAnimData(Escudo)
        .Casco = CascoAnimData(Casco)
        
        .Heading = Heading
        
        .Alpha = 255
        .iTick = 0
        
        'Reset moving stats
        .Moving = 0
        .MoveOffsetX = 0
        .MoveOffsetY = 0
        
        'Update position
        .Pos.x = x
        .Pos.Y = Y
        
        .muerto = Head = CASPER_HEAD Or Head = CASPER_HEAD_CRIMI Or Body = FRAGATA_FANTASMAL
        If .muerto Then .Alpha = 80 Else .Alpha = 255
        'Make active
        .ACTIVE = 1
    End With
    
    'Plot on map
    MapData(x, Y).CharIndex = CharIndex
End Sub

Sub ResetCharInfo(ByVal CharIndex As Integer)
    With charlist(CharIndex)
        .ACTIVE = 0
        .Criminal = 0
        .FxIndex = 0
        .invisible = False
        
#If SeguridadAlkon Then
        Call MI(CualMI).ResetInvisible(CharIndex)
#End If
        
        .Moving = 0
        .muerto = False
        .Alpha = 255
        .iTick = 0
        .ContadorInvi = 0
        .nombre = ""
        .pie = False
        .Pos.x = 0
        .Pos.Y = 0
        .UsandoArma = False
    End With
End Sub

Sub EraseChar(ByVal CharIndex As Integer)
'*****************************************************************
'Erases a character from CharList and map
'*****************************************************************
On Error Resume Next
    charlist(CharIndex).ACTIVE = 0
    
    'Update lastchar
    If CharIndex = LastChar Then
        Do Until charlist(LastChar).ACTIVE = 1
            LastChar = LastChar - 1
            If LastChar = 0 Then Exit Do
        Loop
    End If
    If charlist(CharIndex).Pos.x > 0 And charlist(CharIndex).Pos.Y > 0 Then
    MapData(charlist(CharIndex).Pos.x, charlist(CharIndex).Pos.Y).CharIndex = 0
    
    'Remove char's dialog
    Call Dialogos.RemoveDialog(CharIndex)
    
    Call ResetCharInfo(CharIndex)
    
    'Update NumChars
    NumChars = NumChars - 1
    End If
End Sub

Public Sub InitGrh(ByRef Grh As Grh, ByVal GrhIndex As Integer, Optional ByVal Started As Byte = 2)
'*****************************************************************
'Sets up a grh. MUST be done before rendering
'*****************************************************************
    Grh.GrhIndex = GrhIndex
    If Grh.GrhIndex = 0 Then Exit Sub
    If Started = 2 Then
        If GrhData(Grh.GrhIndex).NumFrames > 1 Then
            Grh.Started = 1
        Else
            Grh.Started = 0
        End If
    Else
        'Make sure the graphic can be started
        If GrhData(Grh.GrhIndex).NumFrames = 1 Then Started = 0
        Grh.Started = Started
    End If
    
    
    If Grh.Started Then
        Grh.Loops = INFINITE_LOOPS
    Else
        Grh.Loops = 0
    End If
    
    Grh.FrameCounter = 1
    Grh.Speed = GrhData(Grh.GrhIndex).Speed
End Sub

Sub MoveCharbyHead(ByVal CharIndex As Integer, ByVal nHeading As E_Heading)
'*****************************************************************
'Starts the movement of a character in nHeading direction
'*****************************************************************
    Dim AddX As Integer
    Dim AddY As Integer
    Dim x As Integer
    Dim Y As Integer
    Dim nX As Integer
    Dim nY As Integer
    Dim tmpInt As Integer
    
    With charlist(CharIndex)
        x = .Pos.x
        Y = .Pos.Y
        
        'Figure out which way to move
        Select Case nHeading
            Case E_Heading.north
                AddY = -1
        
            Case E_Heading.east
                AddX = 1
        
            Case E_Heading.south
                AddY = 1
            
            Case E_Heading.west
                AddX = -1
        End Select
        
        nX = x + AddX
        nY = Y + AddY
        
        If MapData(nX, nY).CharIndex > 0 Then
            tmpInt = MapData(nX, nY).CharIndex
            If charlist(tmpInt).muerto = False Then
                tmpInt = 0
            Else
                charlist(tmpInt).Pos.x = x
                charlist(tmpInt).Pos.Y = Y
                charlist(tmpInt).Heading = InvertHeading(nHeading)
                charlist(tmpInt).MoveOffsetX = 1 * (TilePixelWidth * AddX)
                charlist(tmpInt).MoveOffsetY = 1 * (TilePixelHeight * AddY)
                
                charlist(tmpInt).Moving = 1
                
                charlist(tmpInt).scrollDirectionX = -AddX
                charlist(tmpInt).scrollDirectionY = -AddY
                
                'Si el fantasma soy yo mueve la pantalla
                If tmpInt = UserCharIndex Then Call MoveScreen(charlist(tmpInt).Heading)
            End If
        Else
            tmpInt = 0
        End If
        MapData(nX, nY).CharIndex = CharIndex
        .Pos.x = nX
        .Pos.Y = nY
        MapData(x, Y).CharIndex = tmpInt
        
        .MoveOffsetX = -1 * (TilePixelWidth * AddX)
        .MoveOffsetY = -1 * (TilePixelHeight * AddY)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = AddX
        .scrollDirectionY = AddY
    End With
    
    If UserEstado = 0 Then Call DoPasosFx(CharIndex)
    
    'areas viejos
    'If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
    '    If CharIndex <> UserCharIndex Then
    '        Call EraseChar(CharIndex)
    '    End If
    'End If
End Sub
Public Function InvertHeading(ByVal nHeading As E_Heading) As E_Heading
    Select Case nHeading
        Case E_Heading.east
            InvertHeading = west
        Case E_Heading.west
            InvertHeading = east
        Case E_Heading.south
            InvertHeading = north
        Case E_Heading.north
            InvertHeading = south
    End Select
End Function
Public Sub DoFogataFx()
    Dim location As Position
    
    If bFogata Then
        bFogata = HayFogata(location)
        If Not bFogata Then
            Call Audio.Sound_Stop(FogataBufferIndex)
            FogataBufferIndex = 0
        End If
    Else
        bFogata = HayFogata(location)
        If bFogata And FogataBufferIndex = 0 Then FogataBufferIndex = Audio.Sound_Play(SND_FUEGO, location.x, location.Y, LoopStyle.Enabled)
    End If
End Sub

Public Function EstaPCarea(ByVal CharIndex As Integer) As Boolean
    With charlist(CharIndex).Pos
        EstaPCarea = .x > UserPos.x - MinXBorder And .x < UserPos.x + MinXBorder And .Y > UserPos.Y - MinYBorder And .Y < UserPos.Y + MinYBorder
    End With
End Function

Sub MoveCharbyPos(ByVal CharIndex As Integer, ByVal nX As Integer, ByVal nY As Integer)
On Error Resume Next
    Dim x As Integer
    Dim Y As Integer
    Dim AddX As Integer
    Dim AddY As Integer
    Dim nHeading As E_Heading
    Dim tmpInt As Integer
    
    With charlist(CharIndex)
        x = .Pos.x
        Y = .Pos.Y
        If x > 0 And Y > 0 Then
        
        
        AddX = nX - x
        AddY = nY - Y
        
        If Sgn(AddX) = 1 Then
            nHeading = E_Heading.east
        ElseIf Sgn(AddX) = -1 Then
            nHeading = E_Heading.west
        ElseIf Sgn(AddY) = -1 Then
            nHeading = E_Heading.north
        ElseIf Sgn(AddY) = 1 Then
            nHeading = E_Heading.south
        End If
        
        If MapData(nX, nY).CharIndex > 0 Then
            tmpInt = MapData(nX, nY).CharIndex
            If charlist(tmpInt).muerto = False Then
                tmpInt = 0
            Else
                charlist(tmpInt).Pos.x = x
                charlist(tmpInt).Pos.Y = Y
                charlist(tmpInt).Heading = InvertHeading(nHeading)
                charlist(tmpInt).MoveOffsetX = 1 * (TilePixelWidth * AddX)
                charlist(tmpInt).MoveOffsetY = 1 * (TilePixelHeight * AddY)
                
                charlist(tmpInt).Moving = 1
                
                charlist(tmpInt).scrollDirectionX = -AddX
                charlist(tmpInt).scrollDirectionY = -AddY
                
                'Si el fantasma soy yo mueve la pantalla
                If tmpInt = UserCharIndex Then Call MoveScreen(charlist(tmpInt).Heading)
            End If
        Else
            tmpInt = 0
        End If
        
        
        MapData(x, Y).CharIndex = tmpInt
        
        
        MapData(nX, nY).CharIndex = CharIndex
        
        .Pos.x = nX
        .Pos.Y = nY
        
        .MoveOffsetX = -1 * (TilePixelWidth * AddX)
        .MoveOffsetY = -1 * (TilePixelHeight * AddY)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = Sgn(AddX)
        .scrollDirectionY = Sgn(AddY)
        
        'parche para que no medite cuando camina
        If .FxIndex = FxMeditar.CHICO Or .FxIndex = FxMeditar.GRANDE Or .FxIndex = FxMeditar.MEDIANO Or .FxIndex = FxMeditar.XGRANDE Or .FxIndex = FxMeditar.XXGRANDE Then
            .FxIndex = 0
        End If
        End If
    End With
    
    If Not EstaPCarea(CharIndex) Then Call Dialogos.RemoveDialog(CharIndex)
    
    'If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
    '    Call EraseChar(CharIndex)
    'End If
End Sub

Sub MoveScreen(ByVal nHeading As E_Heading)
'******************************************
'Starts the screen moving in a direction
'******************************************
    Dim x As Integer
    Dim Y As Integer
    Dim tX As Integer
    Dim tY As Integer
    
    'Figure out which way to move
    Select Case nHeading
        Case E_Heading.north
            Y = -1
        
        Case E_Heading.east
            x = 1
        
        Case E_Heading.south
            Y = 1
        
        Case E_Heading.west
            x = -1
    End Select
    
    'Fill temp pos
    tX = UserPos.x + x
    tY = UserPos.Y + Y
    
    'Check to see if its out of bounds
    If tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder Then
        Exit Sub
    Else
        'Start moving... MainLoop does the rest
        AddtoUserPos.x = x
        UserPos.x = tX
        AddtoUserPos.Y = Y
        UserPos.Y = tY
        UserMoving = 1
        
        bTecho = IIf(MapData(UserPos.x, UserPos.Y).Trigger = 1 Or _
                MapData(UserPos.x, UserPos.Y).Trigger = 2 Or _
                MapData(UserPos.x, UserPos.Y).Trigger = 7 Or _
                MapData(UserPos.x, UserPos.Y).Trigger = 4, True, False)
    End If
End Sub

Private Function HayFogata(ByRef location As Position) As Boolean
    Dim j As Long
    Dim k As Long
    
    For j = UserPos.x - 8 To UserPos.x + 8
        For k = UserPos.Y - 6 To UserPos.Y + 6
            If InMapBounds(j, k) Then
                If MapData(j, k).ObjGrh.GrhIndex = GrhFogata Then
                    
                    location.x = j
                    location.Y = k
                    
                    HayFogata = True
                    Exit Function
                End If
            End If
        Next k
    Next j
End Function

Function NextOpenChar() As Integer
'*****************************************************************
'Finds next open char slot in CharList
'*****************************************************************
    Dim loopc As Long
    Dim Dale As Boolean
    
    loopc = 1
    Do While charlist(loopc).ACTIVE And Dale
        loopc = loopc + 1
        Dale = (loopc <= UBound(charlist))
    Loop
    
    NextOpenChar = loopc
End Function

''
' Loads grh data using the new file format.
'
' @return   True if the load was successfull, False otherwise.

Private Function LoadGrhData() As Boolean
On Error GoTo ErrorHandler
    Dim Grh As Long
    Dim Frame As Long
    Dim grhCount As Long
    Dim handle As Integer
    Dim fileVersion As Long
    
    'Open files
    handle = FreeFile()
    
    Open IniPath & GraphicsFile For Binary Access Read As handle
    Seek #1, 1
    
    'Get file version
    Get handle, , fileVersion
    
    'Get number of grhs
    Get handle, , grhCount
    
    'Resize arrays
    ReDim GrhData(1 To grhCount) As GrhData
    
    While Not EOF(handle)
        Get handle, , Grh
        If Grh > 0 Then
        With GrhData(Grh)
            'Get number of frames
            Get handle, , .NumFrames
            If .NumFrames <= 0 Then GoTo ErrorHandler
            
            ReDim .Frames(1 To GrhData(Grh).NumFrames)
            
            If .NumFrames > 1 Then
                'Read a animation GRH set
                For Frame = 1 To .NumFrames
                    Get handle, , .Frames(Frame)
                    If .Frames(Frame) <= 0 Or .Frames(Frame) > grhCount Then
                        GoTo ErrorHandler
                    End If
                Next Frame
                
                Get handle, , .Speed
                
                If .Speed <= 0 Then GoTo ErrorHandler
                
                'Compute width and height
                .pixelHeight = GrhData(.Frames(1)).pixelHeight
                'If .pixelHeight <= 0 Then GoTo ErrorHandler
                
                .pixelWidth = GrhData(.Frames(1)).pixelWidth
                'If .pixelWidth <= 0 Then GoTo ErrorHandler
                
                .TileWidth = GrhData(.Frames(1)).TileWidth
                'If .TileWidth <= 0 Then GoTo ErrorHandler
                
                .TileHeight = GrhData(.Frames(1)).TileHeight
                'If .TileHeight <= 0 Then GoTo ErrorHandler
            Else
                'Read in normal GRH data
                Get handle, , .FileNum
                If .FileNum <= 0 Then GoTo ErrorHandler
                
                Get handle, , GrhData(Grh).sX
                If .sX < 0 Then GoTo ErrorHandler
                
                Get handle, , .sY
                If .sY < 0 Then GoTo ErrorHandler
                
                Get handle, , .pixelWidth
                If .pixelWidth <= 0 Then GoTo ErrorHandler
                
                Get handle, , .pixelHeight
                If .pixelHeight <= 0 Then GoTo ErrorHandler
                
                'Compute width and height
                .TileWidth = .pixelWidth / TilePixelHeight
                .TileHeight = .pixelHeight / TilePixelWidth
                
                .Frames(1) = Grh
            End If
        End With
        End If
    Wend
    
    Close handle
    
    LoadGrhData = True
Exit Function

ErrorHandler:
    LoadGrhData = False
End Function

Function LegalPos(ByVal x As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is legal
'*****************************************************************
    'Limites del mapa
    If x < MinXBorder Or x > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
        Exit Function
    End If
    
    'Tile Bloqueado?
    If MapData(x, Y).Blocked = 1 Then
        Exit Function
    End If
    
    '¿Hay un personaje?
    If MapData(x, Y).CharIndex > 0 Then
        Exit Function
    End If
   
    If UserNavegando <> HayAgua(x, Y) Then
        Exit Function
    End If
    
    LegalPos = True
End Function

Function MoveToLegalPos(ByVal x As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Author: ZaMa
'Last Modify Date: 10/05/2009
'Checks to see if a tile position is legal, including if there is a casper in the tile
'10/05/2009: ZaMa - Now you can't change position with a casper which is in the shore.
'*****************************************************************
    Dim CharIndex As Integer
    
    'Limites del mapa
    If x < MinXBorder Or x > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
        Exit Function
    End If
    
    'Tile Bloqueado?
    If MapData(x, Y).Blocked = 1 Then
        Exit Function
    End If
    
    CharIndex = MapData(x, Y).CharIndex
    '¿Hay un personaje?
    If CharIndex > 0 Then
    
        If MapData(UserPos.x, UserPos.Y).Blocked = 1 Then
            Exit Function
        End If
        
        With charlist(CharIndex)
            ' Si no es casper, no puede pasar
            If .muerto = False Or .nombre = "" Then
                Exit Function
            Else
                ' No puedo intercambiar con un casper que este en la orilla (Lado tierra)
                If HayAgua(UserPos.x, UserPos.Y) Then
                    If Not HayAgua(x, Y) Then Exit Function
                Else
                    ' No puedo intercambiar con un casper que este en la orilla (Lado agua)
                    If HayAgua(x, Y) Then Exit Function
                End If
            End If
        End With
    End If
   
    If UserNavegando <> HayAgua(x, Y) Then
        Exit Function
    End If
    
    MoveToLegalPos = True
End Function

Function InMapBounds(ByVal x As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************
    If x < XMinMapSize Or x > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
        Exit Function
    End If
    
    InMapBounds = True
End Function

Sub DrawGrhIndexLuz(ByVal GrhIndex As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal Center As Byte, ByRef color() As Long)

    With GrhData(GrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                x = x - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
        
        Call Engine_Render_Rectangle(x, Y, .pixelWidth, .pixelHeight, .sX, .sY, .pixelWidth, .pixelHeight, , , , .FileNum, color(0), color(1), color(2), color(3))
    End With
End Sub

Sub DrawGrhIndex(ByVal GrhIndex As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal Center As Byte, ByVal color As Long)

    With GrhData(GrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                x = x - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
        
        Call Engine_Render_Rectangle(x, Y, .pixelWidth, .pixelHeight, .sX, .sY, .pixelWidth, .pixelHeight, , , , .FileNum, color, color, color, color)
    End With
End Sub
Sub DrawGrhLuz(ByRef Grh As Grh, ByVal x As Integer, ByVal Y As Integer, ByVal Center As Byte, ByVal Animate As Single, ByRef color() As Long)
    Dim CurrentGrhIndex As Integer
    
On Error GoTo Error

    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * Animate * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed)
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                
                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If
    End If
    If Grh.GrhIndex > 0 Then
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    With GrhData(CurrentGrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                x = x - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
                
        'If COLOR = -1 Then COLOR = Iluminacion

        Call Engine_Render_Rectangle(x, Y, .pixelWidth, .pixelHeight, .sX, .sY, .pixelWidth, .pixelHeight, , , 0, .FileNum, color(0), color(1), color(2), color(3))
    End With
    End If
Exit Sub

Error:
    If Err.Number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        MsgBox "Ocurrió un error inesperado, por favor comuniquelo a los administradores del juego." & vbCrLf & "Descripción del error: " & _
        vbCrLf & Err.Description, vbExclamation, "[ " & Err.Number & " ] Error"
        End
    End If
End Sub
Sub DrawGrhShadow(ByRef Grh As Grh, ByVal x As Integer, ByVal Y As Integer, ByVal Center As Byte, ByVal Animate As Single, Optional Shadow As Byte = 0, Optional color As Long = -1, Optional ShadowAlpha As Single = 255)
    Dim CurrentGrhIndex As Integer
    
On Error GoTo Error

    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * Animate * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed)
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                
                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    With GrhData(CurrentGrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                x = x - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
                
        'Draw
        'Call BackBufferSurface.BltFast(X, Y, SurfaceDB.Surface(.FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
        If Shadow = 1 Then
            ShadowColor = D3DColorRGBA(0, 0, 0, ShadowAlpha * 100 / 255)
            Call Engine_Render_Rectangle(x, Y, .pixelWidth, .pixelHeight, .sX, .sY, .pixelWidth, .pixelHeight, , , 0, .FileNum, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 1, False)
        ElseIf Shadow = 2 Then
            ShadowColor = D3DColorRGBA(0, 0, 0, ShadowAlpha * 100 / 255)
            Call Engine_Render_Rectangle(x + 10, Y - 16, .pixelWidth, .pixelHeight, .sX, .sY, .pixelWidth, .pixelHeight, , , 0, .FileNum, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 1, False)
        End If
        If color = -1 Then color = Iluminacion
    End With
Exit Sub

Error:
    If Err.Number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        MsgBox "Ocurrió un error inesperado, por favor comuniquelo a los administradores del juego." & vbCrLf & "Descripción del error: " & _
        vbCrLf & Err.Description, vbExclamation, "[ " & Err.Number & " ] Error"
        End
    End If
End Sub
Sub DrawGrhShadowOff(ByRef Grh As Grh, ByVal x As Integer, ByVal Y As Integer, ByVal Center As Byte, ByVal Animate As Single, Optional color As Long = -1)
    Dim CurrentGrhIndex As Integer
    
On Error GoTo Error
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    With GrhData(CurrentGrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                x = x - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
                
        If color = -1 Then color = Iluminacion

        Call Engine_Render_Rectangle(x, Y, .pixelWidth, .pixelHeight, .sX, .sY, .pixelWidth, .pixelHeight, , , 0, .FileNum, color, color, color, color)
    End With
Exit Sub

Error:
    If Err.Number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        MsgBox "Ocurrió un error inesperado, por favor comuniquelo a los administradores del juego." & vbCrLf & "Descripción del error: " & _
        vbCrLf & Err.Description, vbExclamation, "[ " & Err.Number & " ] Error"
        End
    End If
End Sub



Sub DrawGrh(ByRef Grh As Grh, ByVal x As Integer, ByVal Y As Integer, ByVal Center As Byte, ByVal Animate As Single, Optional Shadow As Byte = 0, Optional color As Long = -1, Optional ShadowAlpha As Single = 255)
    Dim CurrentGrhIndex As Integer
    
On Error GoTo Error

    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * Animate * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed)
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                
                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                        If Grh.Loops = 0 Then
                            Grh.Started = 0
                        End If
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    With GrhData(CurrentGrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                x = x - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
                
        'Draw
        'Call BackBufferSurface.BltFast(X, Y, SurfaceDB.Surface(.FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
        If Shadow = 1 Then
            ShadowColor = D3DColorRGBA(0, 0, 0, ShadowAlpha * 100 / 255)
            Call Engine_Render_Rectangle(x, Y, .pixelWidth, .pixelHeight, .sX, .sY, .pixelWidth, .pixelHeight, , , 0, .FileNum, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 1, False)
        ElseIf Shadow = 2 Then
            ShadowColor = D3DColorRGBA(0, 0, 0, ShadowAlpha * 100 / 255)
            Call Engine_Render_Rectangle(x + 10, Y - 16, .pixelWidth, .pixelHeight, .sX, .sY, .pixelWidth, .pixelHeight, , , 0, .FileNum, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 1, False)
        End If
        If color = -1 Then color = Iluminacion

        Call Engine_Render_Rectangle(x, Y, .pixelWidth, .pixelHeight, .sX, .sY, .pixelWidth, .pixelHeight, , , 0, .FileNum, color, color, color, color)
    End With
Exit Sub

Error:
    If Err.Number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        MsgBox "Ocurrió un error inesperado, por favor comuniquelo a los administradores del juego." & vbCrLf & "Descripción del error: " & _
        vbCrLf & Err.Description, vbExclamation, "[ " & Err.Number & " ] Error"
        End
    End If
End Sub

Function GetBitmapDimensions(ByVal BmpFile As String, ByRef bmWidth As Long, ByRef bmHeight As Long)
'*****************************************************************
'Gets the dimensions of a bmp
'*****************************************************************
    Dim BMHeader As BITMAPFILEHEADER
    Dim BINFOHeader As BITMAPINFOHEADER
    
    Open BmpFile For Binary Access Read As #1
    
    Get #1, , BMHeader
    Get #1, , BINFOHeader
    
    Close #1
    
    bmWidth = BINFOHeader.biWidth
    bmHeight = BINFOHeader.biHeight
End Function

Sub DrawGrhtoHdc(ByVal hdc As Long, ByVal GrhIndex As Integer, ByRef SourceRect As RECT, ByRef destRect As RECT)
'*****************************************************************
'Draws a Grh's portion to the given area of any Device Context
'*****************************************************************
    'Call SurfaceDB.Surface(GrhData(GrhIndex).FileNum).BltToDC(hDC, SourceRect, destRect)
Call TransparentBlt(hdc, 0, 0, 32, 32, Inventario.Grafico(GrhData(GrhIndex).FileNum), 0, 0, 32, 32, vbMagenta)
End Sub
Public Sub CargarTile(x As Long, Y As Long)
Dim ByFlags As Byte
Dim Rango As Byte
Dim i As Integer
Dim tmpInt As Integer

Dim Pos As Long

                Pos = (x - 1) * 10 + (Y - 1) * 11000

                ByFlags = DataMap(Pos)
                ByFlags = ByFlags Xor ((x Mod 200) + 55)
                Pos = Pos + 1
            
                If ByFlags = 50 Then
                    MapData(x, Y).Blocked = 1
                Else
                    MapData(x, Y).Blocked = 0
                End If
                MapData(x, Y).Trigger = ByFlags
            
                For i = 1 To 4
                    tmpInt = (DataMap(Pos + 1) And &H7F) * &H100 Or DataMap(Pos) Or -(DataMap(Pos + 1) > &H7F) * &H8000
                    Pos = Pos + 2
                    Select Case i
                        Case 1
                            MapData(x, Y).Graphic(1).GrhIndex = (tmpInt Xor (Y + 301) Xor (x + 721)) - x
                        Case 2
                            MapData(x, Y).Graphic(2).GrhIndex = (tmpInt Xor (Y + 501) Xor (x + 529)) - x
                        Case 3
                            MapData(x, Y).Graphic(3).GrhIndex = (tmpInt Xor (x + 239) Xor (Y + 319)) - x
                        Case 4
                            MapData(x, Y).Graphic(4).GrhIndex = (tmpInt Xor (x + 671) Xor (Y + 129)) - x
                    End Select
                    
                    If MapData(x, Y).Graphic(i).GrhIndex > 0 Then
                        InitGrh MapData(x, Y).Graphic(i), MapData(x, Y).Graphic(i).GrhIndex
                    End If
                Next i
                'Get ArchivoMapa, , Rango
                Rango = DataMap(Pos)
                Pos = Pos + 1
                
                MapData(x, Y).Map = UserMap
                
                MapData(x, Y).Light_Value(0) = Iluminacion
                MapData(x, Y).Light_Value(1) = Iluminacion
                MapData(x, Y).Light_Value(2) = Iluminacion
                MapData(x, Y).Light_Value(3) = Iluminacion
                MapData(x, Y).Hora = Hora
                
                Call Light_Destroy_ToMap(x, Y)
                
                If MapData(x, Y).Graphic(3).GrhIndex < 0 Then
                    Call Light_Create(x, Y, 255, 255, 255, Rango, -MapData(x, Y).Graphic(3).GrhIndex - 1)
                End If
End Sub
Sub RenderScreen(ByVal TileX As Integer, ByVal TileY As Integer, ByVal PixelOffSetX As Single, ByVal PixelOffSetY As Single)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 8/14/2007
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Renders everything to the viewport
'**************************************************************
    Dim Y           As Long     'Keeps track of where on map we are
    Dim x           As Long     'Keeps track of where on map we are
    Dim screenminY  As Integer  'Start Y pos on current screen
    Dim screenmaxY  As Integer  'End Y pos on current screen
    Dim screenminX  As Integer  'Start X pos on current screen
    Dim screenmaxX  As Integer  'End X pos on current screen
    Dim MinY        As Integer  'Start Y pos on current map
    Dim MaxY        As Integer  'End Y pos on current map
    Dim MinX        As Integer  'Start X pos on current map
    Dim MaxX        As Integer  'End X pos on current map
    Dim ScreenX     As Integer  'Keeps track of where to place tile on screen
    Dim ScreenY     As Integer  'Keeps track of where to place tile on screen
    Dim minXOffset  As Integer
    Dim minYOffset  As Integer
    Dim PixelOffsetXTemp As Integer 'For centering grhs
    Dim PixelOffsetYTemp As Integer 'For centering grhs
    Dim tmpInt As Integer
    Dim tmpLong As Long
    Dim SupIndex As Integer
    Dim ByFlags As Byte
    Dim i As Integer
    Dim color As Long
    
    Dim Eliminados As Integer
    Dim Cant As Integer
    
    If UserMap = 0 Then Exit Sub
    
    'Figure out Ends and Starts of screen
    screenminY = TileY - HalfWindowTileHeight
    screenmaxY = TileY + HalfWindowTileHeight
    screenminX = TileX - HalfWindowTileWidth
    screenmaxX = TileX + HalfWindowTileWidth
    
    MinY = screenminY - TileBufferSize
    MaxY = screenmaxY + TileBufferSize
    MinX = screenminX - TileBufferSize
    MaxX = screenmaxX + TileBufferSize
    
    'Make sure mins and maxs are allways in map bounds
    If MinY < XMinMapSize Then
        minYOffset = YMinMapSize - MinY
        MinY = YMinMapSize
    End If
    
    If MaxY > YMaxMapSize Then MaxY = YMaxMapSize
    
    If MinX < XMinMapSize Then
        minXOffset = XMinMapSize - MinX
        MinX = XMinMapSize
    End If
    
    If MaxX > XMaxMapSize Then MaxX = XMaxMapSize
    
    'If we can, we render around the view area to make it smoother
    If screenminY > YMinMapSize Then
        screenminY = screenminY - 1
    Else
        screenminY = 1
        ScreenY = 1
    End If
    
    If screenmaxY < YMaxMapSize Then
        screenmaxY = screenmaxY + 1
    Else
        screenmaxY = YMaxMapSize
    End If
    
    If screenminX > XMinMapSize Then
        screenminX = screenminX - 1
    Else
        screenminX = 1
        ScreenX = 1
    End If
    
    If screenmaxX < XMaxMapSize Then
        screenmaxX = screenmaxX + 1
    Else
        screenmaxX = XMaxMapSize
    End If
    
    'Dim CambioHora As Boolean

    'Cargar mapa
    For Y = MinY - 5 To MaxY + 5
        For x = MinX - 5 To MaxX + 5
            If x > 0 And Y > 0 And x <= XMaxMapSize And Y <= YMaxMapSize Then
                If MapData(x, Y).Map <> UserMap Then
                    Call CargarTile(x, Y)
                End If
                If MapData(x, Y).Hora <> Hora Then
                    For i = 0 To 3
                        MapData(x, Y).Light_Value(i) = Iluminacion
                    Next i
                    MapData(x, Y).Hora = Hora
                    'CambioHora = True
                End If
            End If
        Next x
    Next Y
    
    Light_Render_Area
    
    'Draw floor layer
    For Y = screenminY To screenmaxY
        For x = screenminX To screenmaxX
            
            'Layer 1 **********************************
            Call DrawGrhLuz(MapData(x, Y).Graphic(1), _
                (ScreenX - 1) * TilePixelWidth + PixelOffSetX + TileBufferPixelOffsetX, _
                (ScreenY - 1) * TilePixelHeight + PixelOffSetY + TileBufferPixelOffsetY, _
                 0, 1, MapData(x, Y).Light_Value)
            '******************************************
            ScreenX = ScreenX + 1
        Next x
        
        'Reset ScreenX to original value and increment ScreenY
        ScreenX = ScreenX - x + screenminX
        ScreenY = ScreenY + 1
    Next Y
    
    'Draw floor layer 2
    ScreenY = minYOffset
    For Y = MinY To MaxY
        ScreenX = minXOffset
        For x = MinX To MaxX
            
            'Layer 2 **********************************
            If MapData(x, Y).Graphic(2).GrhIndex <> 0 Then
                Call DrawGrhLuz(MapData(x, Y).Graphic(2), _
                        (ScreenX - 1) * TilePixelWidth + PixelOffSetX, _
                        (ScreenY - 1) * TilePixelHeight + PixelOffSetY, _
                        1, 1, MapData(x, Y).Light_Value)
            End If
            '******************************************
            
            ScreenX = ScreenX + 1
        Next x
        ScreenY = ScreenY + 1
    Next Y
    
    Dim mNPCMuerto As clsNPCMuerto
    
    Eliminados = 0
    Cant = NPCMuertos.Count
    For i = 1 To Cant
        Set mNPCMuerto = NPCMuertos(i - Eliminados)
        Call mNPCMuerto.Update '(TileX, TileY, PixelOffSetX, PixelOffSetY)
        If mNPCMuerto.KillMe Then
            NPCMuertos.Remove (i - Eliminados)
            Eliminados = Eliminados + 1
        End If
    Next i
    
    
    'Draw Transparent Layers
    ScreenY = minYOffset
    For Y = MinY To MaxY
        ScreenX = minXOffset
        For x = MinX To MaxX
            PixelOffsetXTemp = (ScreenX - 1) * TilePixelWidth + PixelOffSetX
            PixelOffsetYTemp = (ScreenY - 1) * TilePixelHeight + PixelOffSetY
            
            With MapData(x, Y)
                'Object Layer **********************************
                If .ObjGrh.GrhIndex <> 0 Then
                    Call DrawGrhLuz(.ObjGrh, _
                            PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, MapData(x, Y).Light_Value)
                End If
                '***********************************************
                
                
                
                If Not .Elemento Is Nothing Then 'Render de Npc Muertos
                    Call .Elemento.Render(PixelOffsetXTemp, PixelOffsetYTemp)
                End If
                
            
                'Char layer ************************************
                If .CharIndex <> 0 Then
                    Call CharRender(charlist(.CharIndex), .CharIndex, PixelOffsetXTemp, PixelOffsetYTemp)
                End If
                
                Call RenderBarcos(x, Y, TileX, TileY, PixelOffSetX, PixelOffSetY)
                
                'Layer 3 *****************************************
                If .Graphic(3).GrhIndex > 0 Then
                    'Draw
                    SupIndex = GrhData(.Graphic(3).GrhIndex).FileNum
                    If ((SupIndex >= 7000 And SupIndex <= 7008) Or (SupIndex >= 1261 And SupIndex <= 1287) Or SupIndex = 648 Or SupIndex = 645) Then
                        If UserPos.x >= x - 4 And UserPos.x <= x + 4 And UserPos.Y >= Y - 7 And UserPos.Y <= Y + 3 Then
                            Call DrawGrh(.Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, 0, D3DColorRGBA(IluRGB.R, IluRGB.G, IluRGB.B, 90))
                        Else
                            Call DrawGrhLuz(.Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, MapData(x, Y).Light_Value)
                        End If
                    Else
                       Call DrawGrh(.Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, 1)
                    End If
                End If
                '*************************************************
            End With
            
            ScreenX = ScreenX + 1
        Next x
        ScreenY = ScreenY + 1
    Next Y
    
    
    Dim mArroja As clsArroja
    Dim Elemento
    For Each Elemento In Arrojas
        Set mArroja = Elemento
        Call mArroja.Render(TileX, TileY, PixelOffSetX, PixelOffSetY)
    Next Elemento
    
    Dim mTooltip As clsToolTip
    
    Eliminados = 0
    Cant = Tooltips.Count
    For i = 1 To Cant
        Set mTooltip = Tooltips(i - Eliminados)
        Call mTooltip.Render(TileX, TileY, PixelOffSetX, PixelOffSetY)
        If mTooltip.Alpha = 0 Then
            Tooltips.Remove (i - Eliminados)
            Eliminados = Eliminados + 1
        End If
    Next i
    
    
    
    If Not bTecho Then
        If bAlpha < 255 Then
            If tTick < (GetTickCount() And &H7FFFFFFF) - 30 Then
                bAlpha = bAlpha + IIf(bAlpha + 8 < 255, 8, 255 - bAlpha)
                ColorTecho = D3DColorRGBA(IluRGB.R, IluRGB.G, IluRGB.B, bAlpha)
                tTick = (GetTickCount() And &H7FFFFFFF)
            End If
        End If
    Else
        If bAlpha > 0 Then
            If tTick < (GetTickCount() And &H7FFFFFFF) - 30 Then
                bAlpha = bAlpha - IIf(bAlpha - 8 > 0, 8, bAlpha)
                ColorTecho = D3DColorRGBA(IluRGB.R, IluRGB.G, IluRGB.B, bAlpha)
                tTick = (GetTickCount() And &H7FFFFFFF)
            End If
        End If
    End If
    'Draw blocked tiles and grid
    ScreenY = minYOffset
    For Y = MinY To MaxY
        ScreenX = minXOffset
        For x = MinX To MaxX
            'Layer 4 **********************************
            If MapData(x, Y).Graphic(4).GrhIndex And bAlpha > 0 Then
                'Draw
                Call DrawGrhIndex(MapData(x, Y).Graphic(4).GrhIndex, _
                    (ScreenX - 1) * TilePixelWidth + PixelOffSetX, _
                    (ScreenY - 1) * TilePixelHeight + PixelOffSetY, _
                    1, ColorTecho)
            End If
                '**********************************
                
            ScreenX = ScreenX + 1
        Next x
        ScreenY = ScreenY + 1
    Next Y
'TODO : Check this!!
    Dim ColorLluvia As Long
    If ZonaActual > 0 Then
        If Zonas(ZonaActual).Terreno <> eTerreno.Dungeon Then
            If bRain Then
                'Figure out what frame to draw
                If llTick < (GetTickCount() And &H7FFFFFFF) - 50 Then
                    iFrameIndex = iFrameIndex + 1
                    If iFrameIndex > 7 Then iFrameIndex = 0
                    llTick = (GetTickCount() And &H7FFFFFFF)
                End If
                ColorLluvia = D3DColorRGBA(IluRGB.R, IluRGB.G, IluRGB.B, 140)
                For Y = 0 To 4
                    For x = 0 To 4
                        Call Engine_Render_Rectangle(LTLluvia(Y), LTLluvia(x), RLluvia(iFrameIndex).Right - RLluvia(iFrameIndex).Left, RLluvia(iFrameIndex).Bottom - RLluvia(iFrameIndex).Top, RLluvia(iFrameIndex).Left, RLluvia(iFrameIndex).Top, RLluvia(iFrameIndex).Right - RLluvia(iFrameIndex).Left, RLluvia(iFrameIndex).Bottom - RLluvia(iFrameIndex).Top, , , , 5556, ColorLluvia, ColorLluvia, ColorLluvia, ColorLluvia)
                    Next x
                Next Y
            End If
        End If
    End If
    
    
    
    Call Dialogos.Render
    Call DibujarCartel
        
    Call DialogosClanes.Draw
    
    If CambioZona > 0 And ZonaActual > 0 Then
        If CambioZona > 300 Then
            tmpInt = 500 - CambioZona
        ElseIf CambioZona < 200 Then
            tmpInt = CambioZona
        Else
            tmpInt = 200
        End If
        If zTick < (GetTickCount() And &H7FFFFFFF) - 50 Then
            CambioZona = CambioZona - 5
            zTick = (GetTickCount() And &H7FFFFFFF)
        End If

        'Mensaje al cambiar de zona
        Call D3DX.DrawText(MainFont, D3DColorRGBA(0, 0, 0, tmpInt), Zonas(ZonaActual).nombre, DDRect(0, 10, 736, 200), DT_CENTER)
        Call D3DX.DrawText(MainFont, D3DColorRGBA(220, 215, 215, tmpInt), Zonas(ZonaActual).nombre, DDRect(5, 15, 736, 200), DT_CENTER)

        If CambioSegura Then
            Call DrawFont(IIf(Zonas(ZonaActual).Segura = 1, "Entraste a una zona segura", "Saliste de una zona segura"), 538, 340, D3DColorRGBA(255, 0, 0, tmpInt))
        End If
    End If
    
    
        If UseMotionBlur Then
        
            AngMareoMuerto = AngMareoMuerto + timerElapsedTime * 0.002
            If AngMareoMuerto >= 6.28318530717959 Then AngMareoMuerto = 0
            
            If GoingHome = 1 Then
                RadioMareoMuerto = RadioMareoMuerto + timerElapsedTime * 0.01
                If RadioMareoMuerto > 50 Then RadioMareoMuerto = 50
            ElseIf GoingHome = 2 Then
                RadioMareoMuerto = RadioMareoMuerto - timerElapsedTime * 0.02
                If RadioMareoMuerto <= 0 Then
                    RadioMareoMuerto = 0
                    GoingHome = 0
                End If
            End If
        
            If FrameUseMotionBlur Then
                FrameUseMotionBlur = False
                With D3DDevice
               
                    'Perform the zooming calculations
                    ' * 1.333... maintains the aspect ratio
                    ' ... / 1024 is to factor in the buffer size
                    BlurTA(0).tu = ZoomLevel + RadioMareoMuerto / 2048 * Sin(AngMareoMuerto) + RadioMareoMuerto / 2048
                    BlurTA(0).tv = ZoomLevel + RadioMareoMuerto / 2048 * Cos(AngMareoMuerto) + RadioMareoMuerto / 2048
                    BlurTA(1).tu = ((ScreenWidth + 1 + Cos(AngMareoMuerto) * RadioMareoMuerto / 2 - RadioMareoMuerto / 2) / 1024) - ZoomLevel
                    BlurTA(1).tv = ZoomLevel + RadioMareoMuerto / 2048 * Sin(AngMareoMuerto) + RadioMareoMuerto / 2048
                    BlurTA(2).tu = ZoomLevel + RadioMareoMuerto / 2048 * Cos(AngMareoMuerto) + RadioMareoMuerto / 2048
                    BlurTA(2).tv = ((ScreenHeight + 1 + Sin(AngMareoMuerto) * RadioMareoMuerto / 2 - RadioMareoMuerto / 2) / 1024) - ZoomLevel
                    BlurTA(3).tu = BlurTA(1).tu
                    BlurTA(3).tv = BlurTA(2).tv
                   
                    'Draw what we have drawn thus far since the last .Clear
                    'LastTexture = -100
                    D3DDevice.EndScene
                    .SetRenderTarget pBackbuffer, Nothing, ByVal 0
                    
                    D3DDevice.BeginScene
                
                    .SetTexture 0, BlurTexture
                    .SetRenderState D3DRS_TEXTUREFACTOR, D3DColorARGB(BlurIntensity, 255, 255, 255)
                    .SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TFACTOR
                    .DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, BlurTA(0), Len(BlurTA(0))
                    .SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
               
                End With
            End If
        End If
    
    
    

        
    If VerMapa Then
        
        '420
        '0.46545454545454545454545454545455
        
        If UserMap = 1 Then
            PosMapX = -Int(UserPos.x * RelacionMiniMapa) + 32 + 368
            PosMapY = -Int(UserPos.Y * RelacionMiniMapa) + 32 + 272
            
            If PosMapX > 0 Then PosMapX = 0
            If PosMapX < -1344 Then PosMapX = -1312
            If PosMapY > 0 Then PosMapY = 0
            If PosMapY < -2273 Then PosMapY = -2273
            
            color = D3DColorRGBA(255, 255, 255, 225)
            
            If PosMapX > -1024 Then 'Dibujo primera columna
                If PosMapY <= 0 And PosMapY > -1024 Then
                    Call Engine_Render_Rectangle(256, 256, _
                                                 736 + IIf(PosMapX < -288, PosMapX + 288, 0), 544 + IIf(PosMapY < -480, PosMapY + 480, 0), _
                                                 -PosMapX, -PosMapY, _
                                                 736 + IIf(PosMapX < -288, PosMapX + 288, 0), 544 + IIf(PosMapY < -480, PosMapY + 480, 0), , , , 14763, color, color, color, color)
                End If
                If PosMapY <= -480 And PosMapY > -2048 Then
                    Call Engine_Render_Rectangle(256, 256 + PosMapY + 1024, _
                                                 736 + IIf(PosMapX < -288, PosMapX + 288, 0), 544 + IIf(PosMapY + 1024 > 0, PosMapY + 1024, 0) + 480, _
                                                 -PosMapX, 0, _
                                                 736 + IIf(PosMapX < -288, PosMapX + 288, 0), 544 + IIf(PosMapY + 1024 > 0, PosMapY + 1024, 0) + 480, , , , 14765, color, color, color, color)
                End If
                If PosMapY <= -1504 Then
                    Call Engine_Render_Rectangle(256, 256 + PosMapY + 2048, _
                                                 736 + IIf(PosMapX < -288, PosMapX + 288, 0), 544 + IIf(PosMapY + 2048 > -480, PosMapY + 2048 + 480, 0), _
                                                 -PosMapX, 0, _
                                                 736 + IIf(PosMapX < -288, PosMapX + 288, 0), 544 + IIf(PosMapY + 2048 > -480, PosMapY + 2048 + 480, 0), , , , 14767, color, color, color, color)
                End If
                            
            End If
            
            If PosMapX < -288 Then 'Dibujo segunda columna
                If PosMapY <= 0 And PosMapY > -1024 Then
                    Call Engine_Render_Rectangle(256 + IIf(PosMapX > -1024, 736 + 288 + PosMapX, 0), 256, _
                                                 736 + IIf(PosMapX < -1024, 0, -1024 - PosMapX), 544 + IIf(PosMapY < -480, PosMapY + 480, 0), _
                                                 IIf(PosMapX < -1024, -PosMapX - 1024, 0), -PosMapY, _
                                                 736 + IIf(PosMapX < -1024, 0, -1024 - PosMapX), 544 + IIf(PosMapY < -480, PosMapY + 480, 0), , , , 14764, color, color, color, color)
                End If
                If PosMapY <= -480 And PosMapY > -2048 Then
                    Call Engine_Render_Rectangle(256 + IIf(PosMapX > -1024, 736 + 288 + PosMapX, 0), 256 + PosMapY + 1024, _
                                                 736 + IIf(PosMapX < -1024, 0, -1024 - PosMapX), 544 + IIf(PosMapY + 1024 > 0, PosMapY + 1024, 0) + 480, _
                                                 IIf(PosMapX < -1024, -PosMapX - 1024, 0), 0, _
                                                 736 + IIf(PosMapX < -1024, 0, -1024 - PosMapX), 544 + IIf(PosMapY + 1024 > 0, PosMapY + 1024, 0) + 480, , , , 14766, color, color, color, color)
                End If
                If PosMapY <= -1504 Then
                    Call Engine_Render_Rectangle(256 + IIf(PosMapX > -1024, 736 + 288 + PosMapX, 0), 256 + PosMapY + 2048, _
                                                 736 + IIf(PosMapX < -1024, 0, -1024 - PosMapX), 544 + IIf(PosMapY + 2048 > -480, PosMapY + 2048 + 480, 0), _
                                                 IIf(PosMapX < -1024, -PosMapX - 1024, 0), 0, _
                                                 736 + IIf(PosMapX < -1024, 0, -1024 - PosMapX), 544 + IIf(PosMapY + 2048 > -480, PosMapY + 2048 + 480, 0), , , , 14768, color, color, color, color)
                End If
            End If
            
            
            'If PosMap <= 210 Then '488
            '    MapaY = 0
            'ElseIf PosMap > 210 And PosMap < 492 Then
            '    MapaY = -(PosMap - 210)
            'Else
            '    MapaY = -282
            'End If
           
            
            'Call Engine_Render_Rectangle(256 + 0, 256 + MapaY, 512, 512, 0, 0, 512, 512, , , , 14404, color, color, color, color)
            'Call Engine_Render_Rectangle(256 + 0, 256 + 512 + MapaY, 512, 186, 0, 0, 512, 186, , , , 14405, color, color, color, color)
           
            color = D3DColorRGBA(255, 255, 255, 255)
            Call Engine_Render_Rectangle(256 + UserPos.x * RelacionMiniMapa - 35 + PosMapX, 256 + UserPos.Y * RelacionMiniMapa - 35 + PosMapY, 4, 4, 0, 0, 4, 4, , , , 1, color, color, color, color)
            
            x = Int((frmMain.MouseX - PosMapX + 32) / RelacionMiniMapa)
            Y = Int((frmMain.MouseY - PosMapY + 32) / RelacionMiniMapa)
            
            If x > 1 And x < 1100 And Y > 1 And Y < 1500 Then
                Call DrawFont("(" & x & "," & Y & ")", frmMain.MouseX + 266, frmMain.MouseY + 266, D3DColorRGBA(255, 255, 255, 200))
                i = BuscarZona(x, Y)
                If i > 0 Then
                    Call DrawFont(Zonas(i).nombre, frmMain.MouseX + 246, frmMain.MouseY + 266 + 13, D3DColorRGBA(255, 255, 255, 200))
                End If
            End If
        ElseIf ZonaActual = 33 Or ZonaActual = 34 Or ZonaActual = 35 Then 'Dungeon Newbie
            color = D3DColorRGBA(255, 255, 255, 190)
            Call Engine_Render_Rectangle(256 + 60, 256 + 3, 512, 512, 0, 0, 512, 512, , , , 14406, color, color, color, color)

            color = D3DColorRGBA(255, 255, 255, 255)
            Call Engine_Render_Rectangle(256 + 60 + (UserPos.x - 571) * 2.21105527638191, 256 + 5 + (UserPos.Y - 311) * 2.21105527638191, 5, 5, 0, 0, 5, 5, , , , 1, color, color, color, color)
        Else
            'Mensaje al cambiar de zona
            Call D3DX.DrawText(MainFont, D3DColorRGBA(0, 0, 0, 200), Zonas(ZonaActual).nombre, DDRect(0, 10, 736, 200), DT_CENTER)
            Call D3DX.DrawText(MainFont, D3DColorRGBA(220, 215, 215, 200), Zonas(ZonaActual).nombre, DDRect(5, 15, 736, 200), DT_CENTER)
        End If
    End If
    
    
D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    
     'Particle_Group_Render MapData(150, 800).particle_group_index, MouseX, MouseY
    
    
    'Dim tmplng As Long
    'Dim tmblng2 As Long
    'ScreenY = minYOffset '- TileBufferSize
    'For y = minY To maxY
    '    ScreenX = minXOffset '- TileBufferSize
    '    For x = minX To maxX
    '        With MapData(x, y)
    '            '*** Start particle effects ***
    '            If MapData(x, y).particle_group_index Then
    '                Particle_Group_Render MapData(x, y).particle_group_index, ScreenX, ScreenY
    '            End If
    '            '*** End particle effects ***
    '        End With
    '        ScreenX = ScreenX + 1
    '    Next x
    '    ScreenY = ScreenY + 1
    'Next y
'Call Engine_Render_Rectangle(frmMain.MouseX, frmMain.MouseY, 128, 128, 0, 256, 128, 128, , , 0, 14332)
                 
D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
 

    If TiempoRetos > 0 Then
        '10 segundos de espera para empezar la ronda
        tmpLong = Abs((GetTickCount() And &H7FFFFFFF) - TiempoRetos)
        tmpInt = 10 - Int(tmpLong / 1000)
        
        If tmpLong < 10000 Then
            Call D3DX.DrawText(MainFontBig, D3DColorRGBA(0, 0, 0, 200), CStr(tmpInt), DDRect(0, 30, 736, 230), DT_CENTER)
            Call D3DX.DrawText(MainFontBig, D3DColorRGBA(220, 215, 215, 200), CStr(tmpInt), DDRect(5, 35, 736, 230), DT_CENTER)
        Else
            'Termino el tiempo de espera que empieze el reto
            TiempoRetos = 0
        End If
    End If
    
    
    If Entrada > 0 Then
        color = D3DColorRGBA(255, 255, 255, Entrada)
        Call Engine_Render_Rectangle(256, 256, 544, 416, 0, 0, 512, 416, , , , 14325, color, color, color, color)
        If zTick2 < (GetTickCount() And &H7FFFFFFF) - 75 Then
            Entrada = Entrada - 15
            zTick2 = (GetTickCount() And &H7FFFFFFF)
        End If
    End If
    
    If PergaminoDireccion > 0 Then
        If PergaminoTick < (GetTickCount() And &H7FFFFFFF) - 20 Then
            If PergaminoDireccion = 1 And AperturaPergamino < 240 Then
                AperturaPergamino = AperturaPergamino + (5 + Sqr(240 - AperturaPergamino) / 2)
                If AperturaPergamino > 240 Then AperturaPergamino = 240
            ElseIf PergaminoDireccion = 2 And AperturaPergamino > 0 Then
                AperturaPergamino = AperturaPergamino - (5 + Sqr(AperturaPergamino) / 2)
                If AperturaPergamino < 0 Then AperturaPergamino = 0
            End If
            PergaminoTick = (GetTickCount() And &H7FFFFFFF)
        End If

    End If
    If AperturaPergamino > 0 Then
        If DateDiff("s", TiempoAbierto, Now) > 10 Then
            PergaminoDireccion = 2
            TiempoAbierto = Now
        End If
        color = D3DColorRGBA(255, 255, 255, AperturaPergamino * 175 / 240)
                
        Call Engine_Render_Rectangle(256 + 10 - 5 + 240 - AperturaPergamino, 256 + 309 + 2, 28, 107, 0, 0, 28, 107, , , , 14687, color, color, color, color)
        Call Engine_Render_Rectangle(256 + 38 - 5 + 240 - AperturaPergamino, 256 + 336 + 2, AperturaPergamino, 74, 240 - AperturaPergamino, 108, AperturaPergamino, 74, , , , 14687, color, color, color, color)
        Call Engine_Render_Rectangle(256 + 517 - 5 - 240 + AperturaPergamino, 256 + 309 + 2, 26, 107, 29, 0, 26, 107, , , , 14687, color, color, color, color)
        Call Engine_Render_Rectangle(256 + 278 - 5, 256 + 335 + 2, AperturaPergamino, 74, 0, 182, AperturaPergamino, 74, , , , 14687, color, color, color, color)
    
        'If AperturaPergamino >= 232 Then
        '    Call Engine_Render_Rectangle(256 + 40, 256 + 335 + 11, 56, 56, 56, 0, 56, 56, , , , 14687, color, color, color, color)
        'ElseIf AperturaPergamino < 232 And AperturaPergamino >= 176 Then
        '    Call Engine_Render_Rectangle(256 + 40 + 232 - AperturaPergamino, 256 + 335 + 11, AperturaPergamino - 176, 56, 56 + 232 - AperturaPergamino, 0, AperturaPergamino - 176, 56, , , , 14687, color, color, color, color)
        'End If
        Call Engine_Render_D3DXTexture(256 + 38 - 5 + 240 - Int(AperturaPergamino), 256 + 342, Int(AperturaPergamino) * 2, 80, 240 - Int(AperturaPergamino), 0, color, pRenderTexture, 0)
    End If

    If FPSFLAG Then Call DrawFont("FPS: " & FPS, 740, 260, D3DColorRGBA(255, 255, 255, 160))
End Sub
Function CalcAlpha(Tiempo As Long, STiempo As Long, MaxAlpha As Byte, Tempo As Single) As Byte
Dim tmpInt As Long

tmpInt = (Tiempo - STiempo) / Tempo
If tmpInt >= 0 Then
CalcAlpha = IIf(tmpInt > MaxAlpha, MaxAlpha, tmpInt)
End If
End Function


Public Function RenderSounds()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/30/2008
'Actualiza todos los sonidos del mapa.
'**************************************************************
If ZonaActual > 0 Then
    If Zonas(ZonaActual).Terreno <> eTerreno.Dungeon Then
        If bRain Then
            If bTecho Then
                If frmMain.IsPlaying <> PlayLoop.plLluviain Then
                    If RainBufferIndex Then _
                        Call Audio.Sound_Stop(RainBufferIndex)
                    RainBufferIndex = Audio.Sound_Play(SND_LLUVIAIN, 0, 0, LoopStyle.Enabled)
                    frmMain.IsPlaying = PlayLoop.plLluviain
                End If
            Else
                If frmMain.IsPlaying <> PlayLoop.plLluviaout Then
                    If RainBufferIndex Then _
                        Call Audio.Sound_Stop(RainBufferIndex)
                    RainBufferIndex = Audio.Sound_Play(SND_LLUVIAOUT, 0, 0, LoopStyle.Enabled)
                    frmMain.IsPlaying = PlayLoop.plLluviaout
                End If
            End If
        End If
    Else
        If frmMain.IsPlaying <> PlayLoop.plNone Then
            Call Audio.Sound_Stop(RainBufferIndex)
            RainBufferIndex = 0
            frmMain.IsPlaying = PlayLoop.plNone
        End If
    End If
    DoFogataFx
    
    
    
End If
End Function

Function HayUserAbajo(ByVal x As Integer, ByVal Y As Integer, ByVal GrhIndex As Integer) As Boolean
    If GrhIndex > 0 Then
        HayUserAbajo = _
            charlist(UserCharIndex).Pos.x >= x - (GrhData(GrhIndex).TileWidth \ 2) _
                And charlist(UserCharIndex).Pos.x <= x + (GrhData(GrhIndex).TileWidth \ 2) _
                And charlist(UserCharIndex).Pos.Y >= Y - (GrhData(GrhIndex).TileHeight - 1) _
                And charlist(UserCharIndex).Pos.Y <= Y
    End If
End Function

Sub LoadGraphics()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero - complete rewrite
'Last Modify Date: 11/03/2006
'Initializes the SurfaceDB and sets up the rain rects
'**************************************************************
    'New surface manager :D
    Call SurfaceDB.Initialize(D3DDevice, D3DX, ClientSetup.bUseVideo, DirRecursos & "Graphics.AO", ClientSetup.byMemory)
    
    'Set up te rain rects
    RLluvia(0).Top = 0:      RLluvia(1).Top = 0:      RLluvia(2).Top = 0:      RLluvia(3).Top = 0
    RLluvia(0).Left = 0:     RLluvia(1).Left = 128:   RLluvia(2).Left = 256:   RLluvia(3).Left = 384
    RLluvia(0).Right = 128:  RLluvia(1).Right = 256:  RLluvia(2).Right = 384:  RLluvia(3).Right = 512
    RLluvia(0).Bottom = 128: RLluvia(1).Bottom = 128: RLluvia(2).Bottom = 128: RLluvia(3).Bottom = 128
    
    RLluvia(4).Top = 128:    RLluvia(5).Top = 128:    RLluvia(6).Top = 128:    RLluvia(7).Top = 128
    RLluvia(4).Left = 0:     RLluvia(5).Left = 128:   RLluvia(6).Left = 256:   RLluvia(7).Left = 384
    RLluvia(4).Right = 128:  RLluvia(5).Right = 256:  RLluvia(6).Right = 384:  RLluvia(7).Right = 512
    RLluvia(4).Bottom = 256: RLluvia(5).Bottom = 256: RLluvia(6).Bottom = 256: RLluvia(7).Bottom = 256
End Sub

Public Function InitTileEngine(ByVal setDisplayFormhWnd As Long, ByVal setMainViewTop As Integer, ByVal setMainViewLeft As Integer, ByVal setTilePixelHeight As Integer, ByVal setTilePixelWidth As Integer, ByVal setWindowTileHeight As Integer, ByVal setWindowTileWidth As Integer, ByVal setTileBufferSize As Integer, ByVal pixelsToScrollPerFrameX As Integer, pixelsToScrollPerFrameY As Integer, ByVal engineSpeed As Single) As Boolean
'***************************************************
'Author: Aaron Perkins
'Last Modification: 08/14/07
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Creates all DX objects and configures the engine to start running.
'***************************************************
    
    IniPath = App.path & "\Init\"
    
    'Fill startup variables
    MainViewTop = setMainViewTop
    MainViewLeft = setMainViewLeft
    TilePixelWidth = setTilePixelWidth
    TilePixelHeight = setTilePixelHeight
    WindowTileHeight = setWindowTileHeight
    WindowTileWidth = setWindowTileWidth
    TileBufferSize = setTileBufferSize
    
    HalfWindowTileHeight = setWindowTileHeight \ 2
    HalfWindowTileWidth = setWindowTileWidth \ 2
    
    'Compute offset in pixels when rendering tile buffer.
    'We diminish by one to get the top-left corner of the tile for rendering.
    TileBufferPixelOffsetX = ((TileBufferSize - 1) * TilePixelWidth)
    TileBufferPixelOffsetY = ((TileBufferSize - 1) * TilePixelHeight)
    
    engineBaseSpeed = engineSpeed
    
    'Set FPS value to 60 for startup
    FPS = 60
    FramesPerSecCounter = 60
    
    MinXBorder = XMinMapSize + (WindowTileWidth \ 2)
    MaxXBorder = XMaxMapSize - (WindowTileWidth \ 2)
    MinYBorder = YMinMapSize + (WindowTileHeight \ 2)
    MaxYBorder = YMaxMapSize - (WindowTileHeight \ 2)
    
    MainViewWidth = TilePixelWidth * WindowTileWidth
    MainViewHeight = TilePixelHeight * WindowTileHeight
    
    'Resize mapdata array
    'ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize, 1 To 2) As MapBlock
    
    'Set intial user position
    UserPos.x = MinXBorder
    UserPos.Y = MinYBorder
    
    'Set scroll pixels per frame
    ScrollPixelsPerFrameX = pixelsToScrollPerFrameX
    ScrollPixelsPerFrameY = pixelsToScrollPerFrameY
    
    'Set the view rect
    With MainViewRect
        .Left = MainViewLeft
        .Top = MainViewTop
        .Right = .Left + MainViewWidth
        .Bottom = .Top + MainViewHeight
    End With
    
    'Set the dest rect
    With MainDestRect
        .Left = TilePixelWidth * TileBufferSize - TilePixelWidth
        .Top = TilePixelHeight * TileBufferSize - TilePixelHeight
        .Right = .Left + MainViewWidth
        .Bottom = .Top + MainViewHeight
    End With
    
    
IniciarD3D

    Call CargarFont
    Call LoadGrhData
    Call CargarCuerpos
    Call CargarCabezas
    Call CargarCascos
    Call CargarFxs
    'Call CargarParticulas
    
    'Call General_Particle_Create(1, 150, 800, -1, 20, -15)
    
   Set TestPart = New clsParticulas
    TestPart.Texture = 14386
        TestPart.ParticleCounts = 35
    TestPart.ReLocate 400, 400
    TestPart.Begin
    
    'Actual = 2
    'Particle_Group_Make Actual, 1, 150, 850, Particula(Actual).VarZ, Particula(Actual).VarX, Particula(Actual).VarY, Particula(Actual).AlphaInicial, Particula(Actual).RedInicial, Particula(Actual).GreenInicial, _
    'Particula(Actual).BlueInicial, Particula(Actual).AlphaFinal, Particula(Actual).RedFinal, Particula(Actual).GreenFinal, Particula(Actual).BlueFinal, Particula(Actual).NumOfParticles, Particula(Actual).gravity, Particula(Actual).Texture, Particula(Actual).Zize, Particula(Actual).Life
    
    
    LTLluvia(0) = 224
    LTLluvia(1) = 352
    LTLluvia(2) = 480
    LTLluvia(3) = 608
    LTLluvia(4) = 736
    
    Call LoadGraphics
    
    InitTileEngine = True
End Function

Sub ShowNextFrame(ByVal DisplayFormTop As Integer, ByVal DisplayFormLeft As Integer, ByVal MouseViewX As Integer, ByVal MouseViewY As Integer)
'***************************************************
'Author: Arron Perkins
'Last Modification: 08/14/07
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Updates the game's model and renders everything.
'***************************************************
    Static OffsetCounterX As Single
    Static OffsetCounterY As Single
    
    
    '****** Set main view rectangle ******
    MainViewRect.Left = (DisplayFormLeft / Screen.TwipsPerPixelX) + MainViewLeft
    MainViewRect.Top = (DisplayFormTop / Screen.TwipsPerPixelY) + MainViewTop
    MainViewRect.Right = MainViewRect.Left + MainViewWidth
    MainViewRect.Bottom = MainViewRect.Top + MainViewHeight
    
    
    If EngineRun Then
        If UserEmbarcado Then
            OffsetCounterX = -BarcoOffSetX
            OffsetCounterY = -BarcoOffSetY
        ElseIf UserMoving Then
            '****** Move screen Left and Right if needed ******
            If AddtoUserPos.x <> 0 Then
                OffsetCounterX = OffsetCounterX - ScrollPixelsPerFrameX * AddtoUserPos.x * timerTicksPerFrame * 1.2
                If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.x) Then
                    OffsetCounterX = 0
                    AddtoUserPos.x = 0
                    UserMoving = False
                End If
            End If
            
            '****** Move screen Up and Down if needed ******
            If AddtoUserPos.Y <> 0 Then
                OffsetCounterY = OffsetCounterY - ScrollPixelsPerFrameY * AddtoUserPos.Y * timerTicksPerFrame * 1.2
                If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.Y) Then
                    OffsetCounterY = 0
                    AddtoUserPos.Y = 0
                    UserMoving = False
                End If
            End If
        End If
        
        'Update mouse position within view area
        Call ConvertCPtoTP(MouseViewX, MouseViewY, MouseTileX, MouseTileY)
           
    If UseMotionBlur Then
        If BlurIntensity < 255 Then
            BlurIntensity = BlurIntensity + (timerElapsedTime * 0.01)
            If BlurIntensity > 255 Then BlurIntensity = 255
        End If
    End If
    If GoingHome Then BlurIntensity = 5
    'Set the motion blur if needed
    If UseMotionBlur Then
        If BlurIntensity < 255 Or ZoomLevel > 0 Then
            FrameUseMotionBlur = True
            D3DDevice.SetRenderTarget BlurSurf, Nothing, ByVal 0
        End If
    End If
        
        
        D3DDevice.BeginScene
        
        'Clear the screen with a solid color (to prevent artifacts)
        If Not FrameUseMotionBlur Then
            D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
        End If
        
        '****** Update screen ******
        If Conectar Then
            Call RenderConectar
        ElseIf UserCiego Then
            Call CleanViewPort
        Else
            Call RenderScreen(UserPos.x - AddtoUserPos.x, UserPos.Y - AddtoUserPos.Y, OffsetCounterX, OffsetCounterY)
        End If
        
    

        'End the rendering (scene)
        D3DDevice.EndScene
                                        
               
        'Flip the backbuffer to the screen
        If Conectar Then
            D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
        Else
            D3DDevice.Present RectJuego, ByVal 0, 0, ByVal 0
        End If

        

    
    
        'Limit FPS to 60 (an easy number higher than monitor's vertical refresh rates)
        'While General_Get_Elapsed_Time2() < 15.5
        '    DoEvents
        'Wend
        
        'timer_ticks_per_frame = General_Get_Elapsed_Time() * 0.029
        
        'FPS update
        If fpsLastCheck + 1000 < (GetTickCount() And &H7FFFFFFF) Then
            FPS = FramesPerSecCounter
            FramesPerSecCounter = 1
            fpsLastCheck = (GetTickCount() And &H7FFFFFFF)
        Else
            FramesPerSecCounter = FramesPerSecCounter + 1
        End If
        
        'Get timing info
        timerElapsedTime = GetElapsedTime()
        If timerElapsedTime <= 0 Then timerElapsedTime = 1
        timerTicksPerFrame = timerElapsedTime * engineBaseSpeed
    End If
End Sub

Public Sub RenderText(ByVal lngXPos As Integer, ByVal lngYPos As Integer, ByRef strText As String, ByVal lngColor As Long)
    If strText <> "" Then
        Call DrawFont(strText, lngXPos, lngYPos, lngColor)
    End If
End Sub

Public Sub RenderTextCentered(ByVal lngXPos As Integer, ByVal lngYPos As Integer, ByRef strText As String, ByVal lngColor As Long)
    If strText <> "" Then
        Call DrawFont(strText, lngXPos, lngYPos, lngColor, True)
    End If
End Sub

Private Function GetElapsedTime() As Single
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the time that past since the last call
'**************************************************************
    Dim start_time As Currency
    Static end_time As Currency
    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        QueryPerformanceFrequency timer_freq
    End If
    
    'Get current time
    Call QueryPerformanceCounter(start_time)
    
    'Calculate elapsed time
    GetElapsedTime = (start_time - end_time) / timer_freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(end_time)
End Function

Public Sub CharRender(ByRef rChar As Char, ByVal CharIndex As Integer, ByVal PixelOffSetX As Integer, ByVal PixelOffSetY As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Draw char's to screen without offcentering them
'***************************************************
    Dim moved As Boolean
    Dim Pos As Integer
    Dim line As String
    Dim color As Long
    Dim VelChar As Single
    Dim ColorPj As Long
    
    With rChar
        If .Moving Then
            If .nombre = "" Then
                VelChar = 0.75
            ElseIf Left(.nombre, 1) = "!" Then
                VelChar = 0.75
            Else
                VelChar = 1.2
            End If
            'If needed, move left and right
            If .scrollDirectionX <> 0 Then
                .MoveOffsetX = .MoveOffsetX + ScrollPixelsPerFrameX * Sgn(.scrollDirectionX) * timerTicksPerFrame * VelChar
                
                'Start animations
'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).Speed > 0 Then _
                    .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionX) = 1 And .MoveOffsetX >= 0) Or _
                        (Sgn(.scrollDirectionX) = -1 And .MoveOffsetX <= 0) Then
                    .MoveOffsetX = 0
                    .scrollDirectionX = 0
                End If
            End If
            
            'If needed, move up and down
            If .scrollDirectionY <> 0 Then
                .MoveOffsetY = .MoveOffsetY + ScrollPixelsPerFrameY * Sgn(.scrollDirectionY) * timerTicksPerFrame * VelChar
                
                'Start animations
'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).Speed > 0 Then _
                    .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionY) = 1 And .MoveOffsetY >= 0) Or _
                        (Sgn(.scrollDirectionY) = -1 And .MoveOffsetY <= 0) Then
                    .MoveOffsetY = 0
                    .scrollDirectionY = 0
                End If
            End If
        End If
        

        
        'If done moving stop animation
        If Not moved And True Then
            'Stop animations
            .Quieto = .Quieto + 1
            If .Quieto >= FPS / 35 Then 'Esto es para que las animacion sean continuas mientras se camine, por ejemplo sin esto el andar del golum se ve feo
            .Quieto = 0
            .Body.Walk(.Heading).Started = 0
            .Body.Walk(.Heading).FrameCounter = 1
            
            If .Arma.WeaponAttack = 0 Then
                .Arma.WeaponWalk(.Heading).Started = 0
                .Arma.WeaponWalk(.Heading).FrameCounter = 1
            Else
                If .Arma.WeaponWalk(.Heading).Started = 0 Then
                    .Arma.WeaponAttack = 0
                    .Arma.WeaponWalk(.Heading).FrameCounter = 1
                End If
            End If
            
            If .Escudo.ShieldAttack = 0 Then
                .Escudo.ShieldWalk(.Heading).Started = 0
                .Escudo.ShieldWalk(.Heading).FrameCounter = 1
            Else
                If .Escudo.ShieldWalk(.Heading).Started = 0 Then
                    .Escudo.ShieldAttack = 0
                    .Escudo.ShieldWalk(.Heading).FrameCounter = 1
                End If
            End If
            End If
            
            .Moving = False
        Else
            .Quieto = 0
        End If
                
        PixelOffSetX = PixelOffSetX + .MoveOffsetX
        PixelOffSetY = PixelOffSetY + .MoveOffsetY
        
        
    
        ColorPj = D3DColorRGBA(IluRGB.R, IluRGB.G, IluRGB.B, .Alpha)
        If .Head.Head(.Heading).GrhIndex Then
                If .invisible Then
                    If CharIndex = UserCharIndex Then
                        .Alpha = 120
                    ElseIf .ContadorInvi > 0 Then
                            If .iTick < (GetTickCount() And &H7FFFFFFF) - 35 Then
                                If .ContadorInvi > 30 And .ContadorInvi <= 60 And .Alpha < 255 Then
                                    .Alpha = .Alpha + 5
                                ElseIf .ContadorInvi <= 30 And .Alpha > 0 Then
                                    .Alpha = .Alpha - 5
                                End If
                                .ContadorInvi = .ContadorInvi - 1
                                .iTick = (GetTickCount() And &H7FFFFFFF)
                            End If
                    Else
                        .ContadorInvi = INTERVALO_INVI
                    End If
                End If
            If .Alpha > 0 Then
                If .priv = 9 Then
                    ColorPj = D3DColorRGBA(10, 10, 10, 255)
                Else
                    ColorPj = D3DColorRGBA(IluRGB.R, IluRGB.G, IluRGB.B, .Alpha)
                End If
                Dim Sombra As Boolean
                If ZonaActual > 0 Then
                    Sombra = .invisible Or .muerto Or Zonas(ZonaActual).Terreno = eTerreno.Dungeon Or .priv = 10
                End If
                'Draw Body
                If .Body.Walk(.Heading).GrhIndex Then _
                    Call DrawGrhShadow(.Body.Walk(.Heading), PixelOffSetX, PixelOffSetY, 1, 0.5, IIf(Sombra, 0, 1), ColorPj)
            
                'Draw Head
                Call DrawGrhShadow(.Head.Head(.Heading), PixelOffSetX + .Body.HeadOffset.x, PixelOffSetY + .Body.HeadOffset.Y, 1, 0, IIf(Sombra, 0, 2), ColorPj)
                    
                'Draw Helmet
                If .Casco.Head(.Heading).GrhIndex Then _
                    Call DrawGrhShadow(.Casco.Head(.Heading), PixelOffSetX + .Body.HeadOffset.x, PixelOffSetY + .Body.HeadOffset.Y, 1, 0, IIf(Sombra, 0, 2), ColorPj)
                    
                'Draw Weapon
                If .Arma.WeaponWalk(.Heading).GrhIndex Then _
                    Call DrawGrhShadow(.Arma.WeaponWalk(.Heading), PixelOffSetX, PixelOffSetY, 1, 0.5, IIf(Sombra, 0, 1), ColorPj)
                    
                'Draw Shield
                If .Escudo.ShieldWalk(.Heading).GrhIndex Then _
                    Call DrawGrhShadow(.Escudo.ShieldWalk(.Heading), PixelOffSetX, PixelOffSetY, 1, 0.5, IIf(Sombra, 0, 1), ColorPj)
                    
                'Draw Body
                If .Body.Walk(.Heading).GrhIndex Then _
                    Call DrawGrhShadowOff(.Body.Walk(.Heading), PixelOffSetX, PixelOffSetY, 1, 0.5, ColorPj)
            
                'Draw Head
                Call DrawGrhShadowOff(.Head.Head(.Heading), PixelOffSetX + .Body.HeadOffset.x, PixelOffSetY + .Body.HeadOffset.Y, 1, 0, ColorPj)
                    
                'Draw Helmet
                If .Casco.Head(.Heading).GrhIndex Then _
                    Call DrawGrhShadowOff(.Casco.Head(.Heading), PixelOffSetX + .Body.HeadOffset.x, PixelOffSetY + .Body.HeadOffset.Y, 1, 0, ColorPj)
                    
                'Draw Weapon
                If .Arma.WeaponWalk(.Heading).GrhIndex Then _
                    Call DrawGrhShadowOff(.Arma.WeaponWalk(.Heading), PixelOffSetX, PixelOffSetY, 1, 0.5, ColorPj)
                    
                'Draw Shield
                If .Escudo.ShieldWalk(.Heading).GrhIndex Then _
                    Call DrawGrhShadowOff(.Escudo.ShieldWalk(.Heading), PixelOffSetX, PixelOffSetY, 1, 0.5, ColorPj)
                
                
                    'Draw name over head
                    If LenB(.nombre) > 0 And Not .invisible And .priv <> 10 Then
                        If Nombres Then
                            Pos = InStr(.nombre, "<")
                            If Pos = 0 Then Pos = Len(.nombre) + 2
                            
                            If .priv = 0 Then
                                If .Criminal Then
                                    color = D3DColorRGBA(ColoresPJ(50).R, ColoresPJ(50).G, ColoresPJ(50).B, 200)
                                Else
                                    color = D3DColorRGBA(ColoresPJ(49).R, ColoresPJ(49).G, ColoresPJ(49).B, 200)
                                End If
                            Else
                                color = D3DColorRGBA(ColoresPJ(.priv).R, ColoresPJ(.priv).G, ColoresPJ(.priv).B, 200)
                            End If
                            
                            'Nick
                            line = Left$(.nombre, Pos - 2)
                            If Left(line, 1) = "!" Then
                                line = Right(line, Len(line) - 1)
                                Pos = Pos - 1
                            End If
                            Call RenderTextCentered(PixelOffSetX + TilePixelWidth \ 2, PixelOffSetY + 30, line, color)
                            
                            'Clan
                            line = mid$(.nombre, Pos)
                            Call RenderTextCentered(PixelOffSetX + TilePixelWidth \ 2, PixelOffSetY + 45, line, color)
                            
                            If .logged Then
                                color = D3DColorRGBA(10, 200, 10, 200)
                                Call RenderTextCentered(PixelOffSetX + TilePixelWidth \ 2, PixelOffSetY + 45, "(Online)", color)
                            End If
                        End If
                    End If
            End If
        Else
            'Draw Body
            If .Body.Walk(.Heading).GrhIndex Then _
                Call DrawGrh(.Body.Walk(.Heading), PixelOffSetX, PixelOffSetY, 1, VelChar, IIf(Sombra, 0, 1), ColorPj)
        End If

        
        'Update dialogs
        Call Dialogos.UpdateDialogPos(PixelOffSetX + .Body.HeadOffset.x + 16, PixelOffSetY + .Body.HeadOffset.Y, CharIndex)
        
        'Draw FX
        If .FxIndex <> 0 Then
D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
            Call DrawGrh(.fX, PixelOffSetX + FxData(.FxIndex).OFFSETX, PixelOffSetY + FxData(.FxIndex).OFFSETY, 1, 1, 0, D3DColorRGBA(255, 255, 255, 170))
D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
            'Check if animation is over
            If .fX.Started = 0 Then _
                .FxIndex = 0
        End If
    End With
End Sub

Public Sub SetCharacterFx(ByVal CharIndex As Integer, ByVal fX As Integer, ByVal Loops As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Sets an FX to the character.
'***************************************************
    With charlist(CharIndex)
        
        
        'If fX > 0 Then
            'If CharIndex = UserCharIndex Then  ' And Not UserMeditar
                .FxIndex = fX
            If fX > 0 Then
                Call InitGrh(.fX, FxData(fX).Animacion)
        
                .fX.Loops = Loops
            End If
            'End If
        'End If
    End With
End Sub

Private Sub CleanViewPort()
'Limpiar
End Sub

