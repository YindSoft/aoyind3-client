Attribute VB_Name = "Mod_General"
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

Public iplst As String

Public bFogata As Boolean

Private lFrameTimer As Long
Public IpServidor As String
Private Type tConsola
    Texto As String
    color As Long
    bold As Byte
    italic As Byte
End Type
Public Consola() As tConsola
Public OffSetConsola As Integer
Public LineasConsola As Integer

Public ArchivoMapa As Integer
Public DataMap(16500000) As Byte

Public Function DirInterface() As String
    DirInterface = App.path & "\" & Config_Inicio.DirGraficos & "\Interface\"
End Function

Public Function DirGraficos() As String
    DirGraficos = App.path & "\" & Config_Inicio.DirGraficos & "\"
End Function

Public Function DirSound() As String
    DirSound = App.path & "\" & Config_Inicio.DirSonidos & "\"
End Function

Public Function DirMidi() As String
    DirMidi = App.path & "\" & Config_Inicio.DirMusica & "\"
End Function


Public Function DirRecursos() As String
    DirRecursos = App.path & "\Recursos\"
End Function

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    'Initialize randomizer
    Randomize Timer
    
    'Generate random number
    RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound
End Function

Sub CargarAnimArmas()
On Error Resume Next
    Dim loopc As Long
    Dim N As Integer
    Dim MisArmas() As tIndiceArma
    N = FreeFile()
    Open App.path & "\init\Armas.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumWeaponAnims
    

    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
    ReDim MisArmas(1 To NumWeaponAnims) As tIndiceArma
    
    For loopc = 1 To NumWeaponAnims
        Get #N, , MisArmas(loopc)
    
        InitGrh WeaponAnimData(loopc).WeaponWalk(1), MisArmas(loopc).Arma(1), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(2), MisArmas(loopc).Arma(2), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(3), MisArmas(loopc).Arma(3), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(4), MisArmas(loopc).Arma(4), 0
    Next loopc
    
    Close #N
    
End Sub



Sub CargarColores()
On Error Resume Next
    Dim archivoC As String
    
    archivoC = App.path & "\init\colores.dat"
    
    If Not FileExist(archivoC, vbArchive) Then
'TODO : Si hay que reinstalar, porque no cierra???
        Call MsgBox("ERROR: no se ha podido cargar los colores. Falta el archivo colores.dat, reinstale el juego", vbCritical + vbOKOnly)
        Exit Sub
    End If
    
    Dim i As Long
    
    For i = 0 To 48 '49 y 50 reservados para ciudadano y criminal
        ColoresPJ(i).R = CByte(GetVar(archivoC, CStr(i), "R"))
        ColoresPJ(i).G = CByte(GetVar(archivoC, CStr(i), "G"))
        ColoresPJ(i).B = CByte(GetVar(archivoC, CStr(i), "B"))
    Next i
    
    ColoresPJ(50).R = CByte(GetVar(archivoC, "CR", "R"))
    ColoresPJ(50).G = CByte(GetVar(archivoC, "CR", "G"))
    ColoresPJ(50).B = CByte(GetVar(archivoC, "CR", "B"))
    ColoresPJ(49).R = CByte(GetVar(archivoC, "CI", "R"))
    ColoresPJ(49).G = CByte(GetVar(archivoC, "CI", "G"))
    ColoresPJ(49).B = CByte(GetVar(archivoC, "CI", "B"))
End Sub

Sub CargarZonas()
On Error Resume Next
    Dim archivoC As String
    
    archivoC = App.path & "\Init\zonas.dat"
    
    If Not FileExist(archivoC, vbArchive) Then
'TODO : Si hay que reinstalar, porque no cierra???
        Call MsgBox("ERROR: no se ha podido cargar las zonas. Falta el archivo zonas.dat, reinstale el juego", vbCritical + vbOKOnly)
        Exit Sub
    End If
    
    
    Dim i As Integer
    Dim e As Integer
    
    NumZonas = GetVar(archivoC, "Config", "Cantidad")
    
    ReDim Zonas(1 To NumZonas)
    For i = 1 To NumZonas
        With Zonas(i)
            .nombre = GetVar(archivoC, "Zona" & CStr(i), "Nombre")
            .Mapa = CByte(GetVar(archivoC, "Zona" & CStr(i), "Mapa"))
            .X1 = CInt(GetVar(archivoC, "Zona" & CStr(i), "X1"))
            .Y1 = CInt(GetVar(archivoC, "Zona" & CStr(i), "Y1"))
            .X2 = CInt(GetVar(archivoC, "Zona" & CStr(i), "X2"))
            .Y2 = CInt(GetVar(archivoC, "Zona" & CStr(i), "Y2"))
            .Segura = CByte(GetVar(archivoC, "Zona" & CStr(i), "Segura"))
            .Acoplar = CByte(Val(GetVar(archivoC, "Zona" & CStr(i), "Acoplar")))
            .Terreno = CByte(Val(GetVar(archivoC, "Zona" & CStr(i), "Terreno")))
            .Musica(1) = Val(GetVar(archivoC, "Zona" & CStr(i), "Musica1"))
            .Musica(2) = Val(GetVar(archivoC, "Zona" & CStr(i), "Musica2"))
            .Musica(3) = Val(GetVar(archivoC, "Zona" & CStr(i), "Musica3"))
            .Musica(4) = Val(GetVar(archivoC, "Zona" & CStr(i), "Musica4"))
            .Musica(5) = Val(GetVar(archivoC, "Zona" & CStr(i), "Musica5"))
            For e = 1 To 5
                If .Musica(e) > 0 Then .CantMusica = .CantMusica + 1
            Next e
        End With
    Next i
End Sub

#If SeguridadAlkon Then
Sub InitMI()
    Dim alternativos As Integer
    Dim CualMITemp As Integer
    
    alternativos = RandomNumber(1, 7368)
    CualMITemp = RandomNumber(1, 1233)
    

    Set MI(CualMITemp) = New clsManagerInvisibles
    Call MI(CualMITemp).Inicializar(alternativos, 10000)
    
    If CualMI <> 0 Then
        Call MI(CualMITemp).CopyFrom(MI(CualMI))
        Set MI(CualMI) = Nothing
    End If
    CualMI = CualMITemp
End Sub
#End If

Sub CargarAnimEscudos()
On Error Resume Next

    Dim loopc As Long
    Dim N As Integer
    Dim MisEscudos() As tIndiceArma

    N = FreeFile()
    Open App.path & "\init\Escudos.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumEscudosAnims
    

    ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
    ReDim MisEscudos(1 To NumEscudosAnims) As tIndiceArma
    
    For loopc = 1 To NumEscudosAnims
        Get #N, , MisEscudos(loopc)
        
        InitGrh ShieldAnimData(loopc).ShieldWalk(1), MisEscudos(loopc).Arma(1), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(2), MisEscudos(loopc).Arma(2), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(3), MisEscudos(loopc).Arma(3), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(4), MisEscudos(loopc).Arma(4), 0
    Next loopc
    
    Close #N
End Sub


'TODO : Never was sure this is really necessary....
'TODO : 08/03/2006 - (AlejoLp) Esto hay que volarlo...
Public Sub RefreshAllChars()
'*****************************************************************
'Goes through the charlist and replots all the characters on the map
'Used to make sure everyone is visible
'*****************************************************************
    Dim loopc As Long
    
    For loopc = 1 To LastChar
        If charlist(loopc).ACTIVE = 1 Then
            MapData(charlist(loopc).Pos.x, charlist(loopc).Pos.Y).CharIndex = loopc
        End If
    Next loopc
End Sub


Function AsciiValidos(ByVal cad As String) As Boolean
    Dim car As Byte
    Dim i As Long
    
    cad = LCase$(cad)
    
    For i = 1 To Len(cad)
        car = Asc(mid$(cad, i, 1))
        
        If ((car < 97 Or car > 122) Or car = Asc("º")) And (car <> 255) And (car <> 32) Then
            Exit Function
        End If
    Next i
    
    AsciiValidos = True
End Function

Function CheckUserData(ByVal checkemail As Boolean) As Boolean
    'Validamos los datos del user
    Dim loopc As Long
    Dim CharAscii As Integer
    
    If checkemail And UserEmail = "" Then
        MessageBox ("Dirección de email invalida")
        Exit Function
    End If
    
    If UserPassword = "" Then
        MessageBox ("Ingrese un password.")
        Exit Function
    End If
    
    For loopc = 1 To Len(UserPassword)
        CharAscii = Asc(mid$(UserPassword, loopc, 1))
        If Not LegalCharacter(CharAscii) Then
            MessageBox ("Password inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
            Exit Function
        End If
    Next loopc
    
    If UserName = "" Then
        MessageBox ("Ingrese un nombre de personaje.")
        Exit Function
    End If
    
    If Len(UserName) > 30 Then
        MessageBox ("El nombre debe tener menos de 30 letras.")
        Exit Function
    End If
    
    For loopc = 1 To Len(UserName)
        CharAscii = Asc(mid$(UserName, loopc, 1))
        If Not LegalCharacter(CharAscii) Then
            MessageBox ("Nombre inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
            Exit Function
        End If
    Next loopc
    
    CheckUserData = True
End Function

Sub UnloadAllForms()
On Error Resume Next

#If SeguridadAlkon Then
    Call UnprotectForm
#End If

    Dim mifrm As Form
    
    For Each mifrm In Forms
        Unload mifrm
    Next
End Sub

Function LegalCharacter(ByVal KeyAscii As Integer) As Boolean
'*****************************************************************
'Only allow characters that are Win 95 filename compatible
'*****************************************************************
    'if backspace allow
    If KeyAscii = 8 Then
        LegalCharacter = True
        Exit Function
    End If
    
    'Only allow space, numbers, letters and special characters
    If KeyAscii < 32 Or KeyAscii = 44 Then
        Exit Function
    End If
    
    If KeyAscii > 126 Then
        Exit Function
    End If
    
    'Check for bad special characters in between
    If KeyAscii = 34 Or KeyAscii = 42 Or KeyAscii = 47 Or KeyAscii = 58 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii = 63 Or KeyAscii = 92 Or KeyAscii = 124 Then
        Exit Function
    End If
    
    'else everything is cool
    LegalCharacter = True
End Function

Sub SetConnected()
'*****************************************************************
'Sets the client to "Connect" mode
'*****************************************************************
    'Set Connected
    Connected = True
    
#If SeguridadAlkon Then
    'Unprotect character creation form
    Call UnprotectForm
#End If
    
    'Unload the connect form
    Unload frmCrearPersonaje
    
    
    
    frmMain.Label8.Caption = UserName
    'Load main form
    frmMain.Visible = True
    
    Conectar = False
    
    Audio.MusicMP3Stop
    
    ZonaActual = 0
    LastZona = ""
    CheckZona
        
    'frmMain.SetRender (True)
    
#If SeguridadAlkon Then
    'Protect the main form
    Call ProtectForm(frmMain)
#End If

End Sub


Sub MoveTo(ByVal Direccion As E_Heading)
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modify Date: 06/28/2008
'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
' 06/03/2006: AlejoLp - Elimine las funciones Move[NSWE] y las converti a esta
' 12/08/2007: Tavo    - Si el usuario esta paralizado no se puede mover.
' 06/28/2008: NicoNZ - Saqué lo que impedía que si el usuario estaba paralizado se ejecute el sub.
'***************************************************
    Dim LegalOk As Boolean
    If Conectar Or UserEmbarcado Then Exit Sub
    If Cartel Then Cartel = False
    
    Select Case Direccion
        Case E_Heading.north
            LegalOk = MoveToLegalPos(UserPos.x, UserPos.Y - 1)
        Case E_Heading.east
            LegalOk = MoveToLegalPos(UserPos.x + 1, UserPos.Y)
        Case E_Heading.south
            LegalOk = MoveToLegalPos(UserPos.x, UserPos.Y + 1)
        Case E_Heading.west
            LegalOk = MoveToLegalPos(UserPos.x - 1, UserPos.Y)
    End Select
    
    If TiempoRetos = 0 Then
        If LegalOk And Not UserParalizado Then
            If Not UserDescansar And Not UserMeditar Then
                Call WriteWalk(Direccion)
                MoveCharbyHead UserCharIndex, Direccion
                MoveScreen Direccion
            End If
        Else
            If charlist(UserCharIndex).Heading <> Direccion Then
                Call WriteChangeHeading(Direccion)
            End If
        End If
    End If
    
    If frmMain.macrotrabajo.Enabled Then frmMain.DesactivarMacroTrabajo
    
    ' Update 3D sounds!
    Call Audio.MoveListener(UserPos.x, UserPos.Y)
    frmMain.Coord.Caption = "(" & UserPos.x & "," & UserPos.Y & ")"
    CheckZona
End Sub

Sub RandomMove()
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modify Date: 06/03/2006
' 06/03/2006: AlejoLp - Ahora utiliza la funcion MoveTo
'***************************************************
    Call MoveTo(RandomNumber(north, west))
End Sub

Private Sub CheckKeys()
'*****************************************************************
'Checks keys and respond
'*****************************************************************
    Static lastMovement As Long
    
    'No input allowed while Argentum is not the active window
    If Not Application.IsAppActive() Then Exit Sub
    
    'No walking when in commerce or banking.
    If Comerciando Then Exit Sub
    
    'If game is paused, abort movement.
    If pausa Then Exit Sub
    
    'Control movement interval (this enforces the 1 step loss when meditating / resting client-side)
    If Abs((GetTickCount() And &H7FFFFFFF) - lastMovement) > 56 Then
        lastMovement = (GetTickCount() And &H7FFFFFFF)
    Else
        Exit Sub
    End If
    
    'Don't allow any these keys during movement..
    If UserMoving = 0 Then
        If Not UserEstupido And Not Conectar Then
            'Move Up
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0 Then
                Call MoveTo(north)
                Exit Sub
            End If
            
            'Move Right
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Then
                Call MoveTo(east)
                Exit Sub
            End If
        
            'Move down
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Then
                Call MoveTo(south)
                Exit Sub
            End If
        
            'Move left
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0 Then
                Call MoveTo(west)
                Exit Sub
            End If
                        
            ' We haven't moved - Update 3D sounds!
            Call Audio.MoveListener(UserPos.x, UserPos.Y)
        Else
            Dim kp As Boolean
            kp = (GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0) Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0
            
            If kp Then
                Call RandomMove
            Else
                ' We haven't moved - Update 3D sounds!
                Call Audio.MoveListener(UserPos.x, UserPos.Y)
            End If
            
            frmMain.Coord.Caption = "(" & UserPos.x & "," & UserPos.Y & ")"
            CheckZona
        End If
    End If
End Sub

'TODO : Si bien nunca estuvo allí, el mapa es algo independiente o a lo sumo dependiente del engine, no va acá!!!

Sub CargarMap(ByVal Map As Integer)
'**************************************************************
'Formato de mapas optimizado para reducir el espacio que ocupan.
'Diseñado y creado por Juan Martín Sotuyo Dodero (Maraxus) (juansotuyo@hotmail.com)
'**************************************************************
    Dim Y As Long
    Dim x As Long
    Dim tempint As Integer
    Dim ByFlags As Byte
    Dim handle As Integer
    
    'If ArchivoMapa > 0 Then
       
    'End If
    
    
    
    ArchivoMapa = FreeFile()
    
    Open DirRecursos & "Mapa" & Map & ".AO" For Binary As ArchivoMapa
        Get #ArchivoMapa, , DataMap
    Close ArchivoMapa
    
For Y = YMinMapSize To YMaxMapSize
        For x = XMinMapSize To XMaxMapSize
            If MapData(x, Y).CharIndex > 0 Then
                Call EraseChar(MapData(x, Y).CharIndex)
            End If
            'Erase OBJs
        MapData(x, Y).ObjGrh.GrhIndex = 0
        Next x
        
Next Y
    
    
    
    Exit Sub
    'Seek handle, 1
            
    'map Header
    'Get handle, , MapInfo.MapVersion
    'Get handle, , MiCabecera
    'Get handle, , tempint
    'Get handle, , tempint
    'Get handle, , tempint
    'Get handle, , tempint
    
    'Load arrays
    For Y = YMinMapSize To IIf(Map = 1, 900, 700) 'YMaxMapSize, 700)
        For x = XMinMapSize To XMaxMapSize
            Get handle, , ByFlags
            
            MapData(x, Y).Blocked = (ByFlags And 1)
            
            Get handle, , MapData(x, Y).Graphic(1).GrhIndex
            InitGrh MapData(x, Y).Graphic(1), MapData(x, Y).Graphic(1).GrhIndex
            
            'Layer 2 used?
            If ByFlags And 2 Then
                Get handle, , MapData(x, Y).Graphic(2).GrhIndex
                InitGrh MapData(x, Y).Graphic(2), MapData(x, Y).Graphic(2).GrhIndex
            Else
                MapData(x, Y).Graphic(2).GrhIndex = 0
            End If
                
            'Layer 3 used?
            If ByFlags And 4 Then
                Get handle, , MapData(x, Y).Graphic(3).GrhIndex
                InitGrh MapData(x, Y).Graphic(3), MapData(x, Y).Graphic(3).GrhIndex
            Else
                MapData(x, Y).Graphic(3).GrhIndex = 0
            End If
                
            'Layer 4 used?
            If ByFlags And 8 Then
                Get handle, , MapData(x, Y).Graphic(4).GrhIndex
                InitGrh MapData(x, Y).Graphic(4), MapData(x, Y).Graphic(4).GrhIndex
            Else
                MapData(x, Y).Graphic(4).GrhIndex = 0
            End If
            
            'Trigger used?
            If ByFlags And 16 Then
                Get handle, , MapData(x, Y).Trigger
            Else
                MapData(x, Y).Trigger = 0
            End If
            
            'Erase NPCs
            If MapData(x, Y).CharIndex > 0 Then
                Call EraseChar(MapData(x, Y).CharIndex)
            End If
            
            If ByFlags And 32 Then Get handle, , tempint
            If ByFlags And 64 Then
                Get handle, , tempint
                Get handle, , tempint
            End If
            If ByFlags And 128 Then
                Get handle, , tempint
                Get handle, , tempint
                Get handle, , tempint
            End If
            'Erase OBJs
            MapData(x, Y).ObjGrh.GrhIndex = 0
        Next x
        If Y Mod 12 = 0 Then
            frmCargando.BProg.Width = frmCargando.BBProg.Width * (0.55 + Y / 3333.333)
            DoEvents
        End If
    Next Y
    
    Close handle

    
End Sub

Sub SwitchMap(ByVal Map As Integer)
'**************************************************************
'Formato de mapas optimizado para reducir el espacio que ocupan.
'Diseñado y creado por Juan Martín Sotuyo Dodero (Maraxus) (juansotuyo@hotmail.com)
'**************************************************************
    UserMap = Map
    CurMap = Map
    
    CargarMap (Map)

End Sub
Sub AddtoRichPicture(ByVal Text As String, Optional ByVal red As Integer = -1, Optional ByVal green As Integer, Optional ByVal blue As Integer, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False, Optional ByVal bCrLf As Boolean = False)
    Dim nId As Long
    Dim AText As String
    Dim Lineas() As String
    Dim i As Integer
    Dim l As Integer
    Dim LastEsp As Integer
    Lineas = Split(Text, vbCrLf)
    For l = 0 To UBound(Lineas)
    Text = Lineas(l)
    nId = LineasConsola + 1
    If nId = 601 Then
        For i = 0 To 500
            Consola(i) = Consola(i + 100)
        Next i
        nId = 501
        If OffSetConsola > 101 Then OffSetConsola = OffSetConsola - 100
    End If
    LineasConsola = nId
    frmMain.pConsola.FontBold = bold
    frmMain.pConsola.FontItalic = italic
    Consola(nId).Texto = Text
    Consola(nId).color = RGB(red, green, blue)
    Consola(nId).bold = bold
    Consola(nId).italic = italic
    If LineasConsola > 6 Then
        OffSetConsola = LineasConsola - 6
        frmMain.BarritaConsola.Top = 68
    End If
    If frmMain.pConsola.TextWidth(Text) > frmMain.pConsola.Width Then
        LastEsp = 0
        For i = 1 To Len(Text)
            If mid(Text, i, 1) = " " Then LastEsp = i
            If frmMain.pConsola.TextWidth(Left$(Text, i)) > frmMain.pConsola.Width Then Exit For
        Next i
        If LastEsp = 0 Then LastEsp = i
        AText = Right$(Text, Len(Text) - LastEsp)
        Text = Left$(Text, LastEsp)
        Consola(nId).Texto = Text
        Call AddtoRichPicture(AText, red, green, blue, bold, italic)
    Else
        frmMain.ReDrawConsola
    End If
    Next l
End Sub
Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Byte) As String
'*****************************************************************
'Gets a field from a delimited string
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/15/2004
'*****************************************************************
    Dim i As Long
    Dim lastPos As Long
    Dim CurrentPos As Long
    Dim delimiter As String * 1
    
    delimiter = Chr$(SepASCII)
    
    For i = 1 To Pos
        lastPos = CurrentPos
        CurrentPos = InStr(lastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i
    
    If CurrentPos = 0 Then
        ReadField = mid$(Text, lastPos + 1, Len(Text) - lastPos)
    Else
        ReadField = mid$(Text, lastPos + 1, CurrentPos - lastPos - 1)
    End If
End Function

Function FieldCount(ByRef Text As String, ByVal SepASCII As Byte) As Long
'*****************************************************************
'Gets the number of fields in a delimited string
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 07/29/2007
'*****************************************************************
    Dim Count As Long
    Dim curPos As Long
    Dim delimiter As String * 1
    
    If LenB(Text) = 0 Then Exit Function
    
    delimiter = Chr$(SepASCII)
    
    curPos = 0
    
    Do
        curPos = InStr(curPos + 1, Text, delimiter)
        Count = Count + 1
    Loop While curPos <> 0
    
    FieldCount = Count
End Function

Function FileExist(ByVal file As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (Dir$(file, FileType) <> "")
End Function



Public Function IsIp(ByVal Ip As String) As Boolean
    Dim i As Long
    
    For i = 1 To UBound(ServersLst)
        If ServersLst(i).Ip = Ip Then
            IsIp = True
            Exit Function
        End If
    Next i
End Function


Sub Main()

    'Load config file
    If FileExist(App.path & "\init\Inicio.con", vbNormal) Then
        Config_Inicio = LeerGameIni()
    End If
    

    Call InitDebug

    'Load ao.dat config file
    If FileExist(App.path & "\init\ao.dat", vbArchive) Then
        Call LoadClientSetup
        If ClientSetup.bDinamic Then
            Set SurfaceDB = New clsSurfaceManDyn
        Else
            Set SurfaceDB = New clsSurfaceManStatic
        End If
    Else
        'Use dynamic by default
        Set SurfaceDB = New clsSurfaceManDyn
    End If
    
    If FindPreviousInstance Then
        Call MsgBox("AoYind ya esta corriendo! No es posible correr otra instancia del juego. Haga click en Aceptar para salir.", vbApplicationModal + vbInformation + vbOKOnly, "Error al ejecutar")
        'End
    End If
    
    Call LeerLineaComandos
    'Read command line. Do it AFTER config file is loaded to prevent this from
    'canceling the effects of "/nores" option.
   
    ReDim SurfaceSize(15000)
    ReDim Consola(600)

    
    
    'usaremos esto para ayudar en los parches
    Call SaveSetting("ArgentumOnlineCliente", "Init", "Path", App.path & "\")
    
    ChDrive App.path
    ChDir App.path

#If SeguridadAlkon Then
    'Obtener el HushMD5
    Dim fMD5HushYo As String * 32
    
    fMD5HushYo = MD5.GetMD5File(App.path & "\" & App.EXEName & ".exe")
    Call MD5.MD5Reset
    MD5HushYo = txtOffset(hexMd52Asc(fMD5HushYo), 55)
    
    Debug.Print fMD5HushYo
#Else
    MD5HushYo = "0123456789abcdef"  'We aren't using a real MD5
#End If
    
    tipf = Config_Inicio.tip
    
    'Set resolution BEFORE the loading form is displayed, therefore it will be centered.
    Call Resolution.SetResolution
    
    Set picMouseIcon = LoadPicture(DirRecursos & "Hand.ico")
    
    frmCargando.Show
    frmCargando.Refresh
    
    frmCargando.BProg.Width = frmCargando.BBProg.Width * 0.05
    DoEvents

    frmCargando.BProg.Width = frmCargando.BBProg.Width * 0.2
    DoEvents
    
    
'TODO : esto de ServerRecibidos no se podría sacar???
    ServersRecibidos = True
    

    Call InicializarNombres
    
    ' Initialize FONTTYPES
    Call Protocol.InitFonts
       
    frmCargando.BProg.Width = frmCargando.BBProg.Width * 0.25
    DoEvents
    
    If Not InitTileEngine(frmMain.hwnd, 160, 7, 32, 32, 17, 23, 9, 8, 8, 0.018) Then
        Call CloseClient
    End If
    
    frmCargando.BProg.Width = frmCargando.BBProg.Width * 0.4
    DoEvents
    

    UserMap = 0
    
    Call CargarAnimArmas
    Call CargarAnimEscudos
    Call CargarColores
    Call CargarZonas
    Call CargarPasos
    Call CargarTutorial
    
    frmCargando.BProg.Width = frmCargando.BBProg.Width * 0.45
    DoEvents

    'Inicializamos el sonido
    Call Audio.Initialize(dX, frmMain.hwnd, App.path & "\" & Config_Inicio.DirSonidos & "\", App.path & "\" & Config_Inicio.DirMusica & "\")
    'Enable / Disable audio
    Audio.MusicActivated = Not ClientSetup.bNoMusic
    Audio.SoundActivated = Not ClientSetup.bNoSound
    Audio.SoundEffectsActivated = Not ClientSetup.bNoSoundEffects
    
    'Audio.SoundVolume = 100
    
    'Inicializamos el inventario gráfico
    Call Inventario.Initialize(frmMain.picInv, 32, 51, MAX_INVENTORY_SLOTS, True)
    
    frmCargando.BProg.Width = frmCargando.BBProg.Width * 0.55
    DoEvents
    
    'Call CargarMap(1)
    'Call CargarMap(2)
    
    frmCargando.BProg.Width = frmCargando.BBProg.Width * 1
    DoEvents
    
#If SeguridadAlkon Then
    CualMI = 0
    Call InitMI
#End If
    

    Unload frmCargando
    
    'Call Audio.PlayMIDI(MIdi_Inicio & ".mid")

    Set frmMain.Client = New CSocketMaster

    frmMain.SetRender (True)
    frmMain.Show
    
    'Inicialización de variables globales
    PrimeraVez = True
    prgRun = True
    pausa = False
    
    'Set the intervals of timers
    Call MainTimer.SetInterval(TimersIndex.Attack, INT_ATTACK)
    Call MainTimer.SetInterval(TimersIndex.Work, INT_WORK)
    Call MainTimer.SetInterval(TimersIndex.UseItemWithU, INT_USEITEMU)
    Call MainTimer.SetInterval(TimersIndex.UseItemWithDblClick, INT_USEITEMDCK)
    Call MainTimer.SetInterval(TimersIndex.SendRPU, INT_SENTRPU)
    Call MainTimer.SetInterval(TimersIndex.CastSpell, INT_CAST_SPELL)
    Call MainTimer.SetInterval(TimersIndex.Arrows, INT_ARROWS)
    Call MainTimer.SetInterval(TimersIndex.CastAttack, INT_CAST_ATTACK)
    Call MainTimer.SetInterval(TimersIndex.Hide, INT_HIDE)
    Call MainTimer.SetInterval(TimersIndex.Buy, INT_BUY)
    
    frmMain.macrotrabajo.Interval = INT_MACRO_TRABAJO
    frmMain.macrotrabajo.Enabled = False
    
   'Init timers
    Call MainTimer.Start(TimersIndex.Attack)
    Call MainTimer.Start(TimersIndex.Work)
    Call MainTimer.Start(TimersIndex.UseItemWithU)
    Call MainTimer.Start(TimersIndex.UseItemWithDblClick)
    Call MainTimer.Start(TimersIndex.SendRPU)
    Call MainTimer.Start(TimersIndex.CastSpell)
    Call MainTimer.Start(TimersIndex.Arrows)
    Call MainTimer.Start(TimersIndex.CastAttack)
    Call MainTimer.Start(TimersIndex.Hide)
    Call MainTimer.Start(TimersIndex.Buy)
    
    'Set the dialog's font
    Dialogos.font = frmMain.font
    DialogosClanes.font = frmMain.font
    
    
    ' Load the form for screenshots
    Call Load(frmScreenshots)
    
    If False Then 'ipx
        IpServidor = "ip publica"
    Else
        IpServidor = frmMain.Client.LocalIP 'localhost
    End If
    
    'IpServidor = "aoyind3.no-ip.org" ' frmMain.Client.LocalIP ' "aoyind3.no-ip.org" ' "javiercasa.no-ip.info"  "aoyind3.no-ip.org" ' 'frmMain.Client.LocalIP '
    
    
    Call Audio.MusicMP3Play("9.mp3")
    
    Call InitBarcos
    
    
    GTCInicial = (GetTickCount() And &H7FFFFFFF)
    
    Conectar = True
    EngineRun = True
    
    Nombres = True

    Do While prgRun
        'Sólo dibujamos si la ventana no está minimizada
        
        
        If False Then
        
            If iTickMidi < (GetTickCount() And &H7FFFFFFF) - 200 Then

                If VolumenCambio > 0 Then
                    VolumenCambio = VolumenCambio - 5
                    If Not Audio.PlayingMusic Then VolumenCambio = 0
                    If VolumenCambio > 40 Then
                        Audio.MusicVolume = (VolumenCambio)
                    Else
                        VolumenCambio = 0
                        Call Audio.StopMidi
                        Call Audio.PlayMIDI(MidiCambio & ".mid")
                        Audio.MusicVolume = (100)
                    End If
                    
                    iTickMidi = (GetTickCount() And &H7FFFFFFF)
                End If
                
            End If
        

        End If
        
        
        
        If frmMain.WindowState <> 1 And frmMain.Visible Then
            Call CalcularBarcos
            Call ShowNextFrame(frmMain.Top, frmMain.Left, frmMain.MouseX, frmMain.MouseY)
            
            'Play ambient sounds
            Call RenderSounds
            
            Call CheckKeys
        Else
            Call CalcularBarcos
        End If
        'FPS Counter - mostramos las FPS
        If Abs((GetTickCount() And &H7FFFFFFF) - lFrameTimer) >= 1000 Then
            lFrameTimer = (GetTickCount() And &H7FFFFFFF)
        End If
        
#If SeguridadAlkon Then
        Call CheckSecurity
#End If
        
        ' If there is anything to be sent, we send it
        Call FlushBuffer
        
        DoEvents
    Loop
    
    Call CloseClient
End Sub

Sub WriteVar(ByVal file As String, ByVal Main As String, ByVal Var As String, ByVal Value As String)
'*****************************************************************
'Writes a var to a text file
'*****************************************************************
    writeprivateprofilestring Main, Var, Value, file
End Sub

Function GetVar(ByVal file As String, ByVal Main As String, ByVal Var As String) As String
'*****************************************************************
'Gets a Var from a text file
'*****************************************************************
    Dim sSpaces As String ' This will hold the input that the program will retrieve
    
    sSpaces = Space$(100) ' This tells the computer how long the longest string can be. If you want, you can change the number 100 to any number you wish
    
    getprivateprofilestring Main, Var, vbNullString, sSpaces, Len(sSpaces), file
    
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

'[CODE 002]:MatuX
'
'  Función para chequear el email
'
'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba y evitar un chequeo innecesario
Public Function CheckMailString(ByVal sString As String) As Boolean
On Error GoTo errHnd
    Dim lPos  As Long
    Dim lX    As Long
    Dim iAsc  As Integer
    
    '1er test: Busca un simbolo @
    lPos = InStr(sString, "@")
    If (lPos <> 0) Then
        '2do test: Busca un simbolo . después de @ + 1
        If Not (InStr(lPos, sString, ".", vbBinaryCompare) > lPos + 1) Then _
            Exit Function
        
        '3er test: Recorre todos los caracteres y los valída
        For lX = 0 To Len(sString) - 1
            If Not (lX = (lPos - 1)) Then   'No chequeamos la '@'
                iAsc = Asc(mid$(sString, (lX + 1), 1))
                If Not CMSValidateChar_(iAsc) Then _
                    Exit Function
            End If
        Next lX
        
        'Finale
        CheckMailString = True
    End If
errHnd:
End Function

'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba
Private Function CMSValidateChar_(ByVal iAsc As Integer) As Boolean
    CMSValidateChar_ = (iAsc >= 48 And iAsc <= 57) Or _
                        (iAsc >= 65 And iAsc <= 90) Or _
                        (iAsc >= 97 And iAsc <= 122) Or _
                        (iAsc = 95) Or (iAsc = 45) Or (iAsc = 46)
End Function

'TODO : como todo lo relativo a mapas, no tiene nada que hacer acá....
Function HayAgua(ByVal x As Integer, ByVal Y As Integer) As Boolean
    HayAgua = ((MapData(x, Y).Graphic(1).GrhIndex >= 1505 And MapData(x, Y).Graphic(1).GrhIndex <= 1520) Or _
            (MapData(x, Y).Graphic(1).GrhIndex >= 5665 And MapData(x, Y).Graphic(1).GrhIndex <= 5680) Or _
            (MapData(x, Y).Graphic(1).GrhIndex >= 13547 And MapData(x, Y).Graphic(1).GrhIndex <= 13562)) And _
                MapData(x, Y).Graphic(2).GrhIndex = 0
                
End Function

Public Sub ShowSendTxt()
    If Not frmCantidad.Visible Then
        frmMain.SendTxt.Visible = True
        frmMain.SendTxt.SetFocus
    End If
End Sub

Public Sub ShowSendCMSGTxt()
    If Not frmCantidad.Visible Then
        frmMain.SendCMSTXT.Visible = True
        frmMain.SendCMSTXT.SetFocus
    End If
End Sub

''
' Checks the command line parameters, if you are running Ao with /nores command and checks the AoUpdate parameters
'
'

Public Sub LeerLineaComandos()
'*************************************************
'Author: Unknown
'Last modified: 25/11/2008 (BrianPr)
'
'*************************************************
    Dim T() As String
    Dim i As Long
    
    Dim UpToDate As Boolean
    Dim Patch As String
    
    'Parseo los comandos
    T = Split(Command, " ")
    For i = LBound(T) To UBound(T)
        Select Case UCase$(T(i))
            Case "/NORES" 'no cambiar la resolucion
                NoRes = True
            Case "/UPTODATE"
                UpToDate = True
        End Select
    Next i
    NoRes = True
    UpToDate = True
    Call AoUpdate(UpToDate, NoRes)
End Sub

''
' Runs AoUpdate if we haven't updated yet, patches aoupdate and runs Client normally if we are updated.
'
' @param UpToDate Specifies if we have checked for updates or not
' @param NoREs Specifies if we have to set nores arg when running the client once again (if the AoUpdate is executed).

Private Sub AoUpdate(ByVal UpToDate_ As Boolean, ByVal NoRes_ As Boolean)
'*************************************************
'Author: BrianPr
'Created: 25/11/2008
'Last modified: 25/11/2008
'
'*************************************************
    Dim extraArgs As String
    Dim Reintentos As Integer

    If Not UpToDate_ Then
        'No recibe update, ejecutar AU
        'Ejecuto el AoUpdate, sino me voy
        If Dir(App.path & "\AoUpdate.exe", vbArchive) = vbNullString Then
            MsgBox "No se encuentra el archivo de actualización AoUpdate.exe por favor descarguelo y vuelva a intentar", vbCritical
            End
        Else
Reintentar:
On Error GoTo Error
            'FileCopy App.path & "\AoUpdate.exe", App.path & "\AoUpdateTMP.exe"
            If NoRes_ Then
                extraArgs = " /nores"
            End If
            
            Call ShellExecute(0, "Open", App.path & "\AoUpdate.exe", App.EXEName & ".exe", App.path, SW_SHOWNORMAL)
            'Call Shell(App.path & "\AoUpdateTMP.exe", App.EXEName & ".exe")
            End
            Exit Sub
        End If
    Else
        If FileExist(App.path & "\AoUpdateTMP.exe", vbArchive) Then Kill App.path & "\AoUpdateTMP.exe"
    End If
Exit Sub

Error:
    If Err.Number = 75 Then 'Si el archivo AoUpdateTMP.exe está en uso, entonces esperamos 5 ms y volvemos a intentarlo hasta que nos deje.
        Reintentos = Reintentos + 1
        If Reintentos = 3 Then
            Call MsgBox("El proceso AoUpdateTMP.exe se encuentra abierto o protegido y no es posible reemplazarlo. Cierre el proceso y vuelva a ejecutar el juego.", vbError)
            End
        Else
        Sleep 500
        GoTo Reintentar:
        End If
        
    Else
        MsgBox Err.Description & vbCrLf, vbInformation, "[ " & Err.Number & " ]" & " Error "
        End
    End If
End Sub

Private Sub LoadClientSetup()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 24/06/2006
'
'**************************************************************
    Dim fHandle As Integer
    
    fHandle = FreeFile
    Open App.path & "\init\ao.dat" For Binary Access Read Lock Write As fHandle
        Get fHandle, , ClientSetup
    Close fHandle
    
    NoRes = ClientSetup.bNoRes
    
    ClientSetup.WinSock = True
    
    GraphicsFile = "Graficos.ind"
End Sub
Private Sub SaveClientSetup()
'**************************************************************
'Author: Torres Patricio (Pato)
'Last Modify Date: 03/11/10
'
'**************************************************************
    Dim fHandle As Integer
    
    fHandle = FreeFile
    
    ClientSetup.bNoMusic = Not Audio.MusicActivated
    ClientSetup.bNoSound = Not Audio.SoundActivated
    ClientSetup.bNoSoundEffects = Not Audio.SoundEffectsActivated
    'ClientSetup.bGuildNews = Not ClientSetup.bGuildNews
    'ClientSetup.bGldMsgConsole = Not DialogosClanes.Activo
    'ClientSetup.bCantMsgs = DialogosClanes.CantidadDialogos
    
    Open App.path & "\Init\AO.dat" For Binary As fHandle
        Put fHandle, , ClientSetup
    Close fHandle
End Sub
Private Sub InicializarNombres()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'Inicializa los nombres de razas, ciudades, clases, skills, atributos, etc.
'**************************************************************
    Ciudades(eCiudad.cUllathorpe) = "Ullathorpe"
    Ciudades(eCiudad.cNix) = "Nix"
    Ciudades(eCiudad.cBanderbill) = "Banderbill"
    Ciudades(eCiudad.cArkhein) = "Arkhein"
    Ciudades(eCiudad.cArghal) = "Arghâl"
    Ciudades(eCiudad.cLindos) = "Lindos"
    
    ListaRazas(eRaza.Humano) = "Humano"
    ListaRazas(eRaza.Elfo) = "Elfo"
    ListaRazas(eRaza.ElfoOscuro) = "Elfo Oscuro"
    ListaRazas(eRaza.Gnomo) = "Gnomo"
    ListaRazas(eRaza.Enano) = "Enano"

    ListaClases(eClass.Mage) = "Mago"
    ListaClases(eClass.Cleric) = "Clerigo"
    ListaClases(eClass.Warrior) = "Guerrero"
    ListaClases(eClass.Assasin) = "Asesino"
    ListaClases(eClass.Thief) = "Ladron"
    ListaClases(eClass.Bard) = "Bardo"
    ListaClases(eClass.Druid) = "Druida"
    ListaClases(eClass.Bandit) = "Bandido"
    ListaClases(eClass.Paladin) = "Paladin"
    ListaClases(eClass.Hunter) = "Cazador"
    ListaClases(eClass.Worker) = "Trabajador"
    ListaClases(eClass.Pirat) = "Pirata"
    
    SkillsNames(eSkill.Magia) = "Magia"
    SkillsNames(eSkill.Robar) = "Robar"
    SkillsNames(eSkill.Tacticas) = "Evasión en combate"
    SkillsNames(eSkill.Armas) = "Combate con armas"
    SkillsNames(eSkill.Meditar) = "Meditar"
    SkillsNames(eSkill.Apuñalar) = "Apuñalar"
    SkillsNames(eSkill.Ocultarse) = "Ocultarse"
    SkillsNames(eSkill.Supervivencia) = "Supervivencia"
    SkillsNames(eSkill.Talar) = "Talar árboles"
    SkillsNames(eSkill.Comerciar) = "Comercio"
    SkillsNames(eSkill.Defensa) = "Defensa con escudos"
    SkillsNames(eSkill.Pesca) = "Pesca"
    SkillsNames(eSkill.Mineria) = "Mineria"
    SkillsNames(eSkill.Carpinteria) = "Carpinteria"
    SkillsNames(eSkill.Herreria) = "Herreria"
    SkillsNames(eSkill.Liderazgo) = "Liderazgo"
    SkillsNames(eSkill.Domar) = "Domar animales"
    SkillsNames(eSkill.Proyectiles) = "Combate a distancia"
    SkillsNames(eSkill.Wrestling) = "Combate sin armas"
    SkillsNames(eSkill.Navegacion) = "Navegacion"

    AtributosNames(eAtributos.Fuerza) = "Fuerza"
    AtributosNames(eAtributos.Agilidad) = "Agilidad"
    AtributosNames(eAtributos.Inteligencia) = "Inteligencia"
    AtributosNames(eAtributos.Carisma) = "Carisma"
    AtributosNames(eAtributos.Constitucion) = "Constitucion"
End Sub

''
' Removes all text from the console and dialogs

Public Sub CleanDialogs()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'Removes all text from the console and dialogs
'**************************************************************
    'Clean console and dialogs
    LineasConsola = 0
    OffSetConsola = 0
    ZonaActual = 0
    CambioZona = 0
    
    Call DialogosClanes.RemoveDialogs
    
    Call Dialogos.RemoveAllDialogs
End Sub
Public Sub CloseSock()
If Not ClientSetup.WinSock Then
    frmMain.Client.CloseSck
Else
    frmMain.WSock.Close
End If
End Sub
Public Sub CloseClient()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 8/14/2007
'Frees all used resources, cleans up and leaves
'**************************************************************
    ' Allow new instances of the client to be opened
    Call PrevInstance.ReleaseInstance
    
    CloseSock
    
    EngineRun = False
    frmCargando.Show

    Call SaveClientSetup

    Call Resolution.ResetResolution
    
    'Stop tile engine
    Call LiberarObjetosDX
    
    'Destruimos los objetos públicos creados
    Set CustomMessages = Nothing
    Set CustomKeys = Nothing
    Set SurfaceDB = Nothing
    Set Dialogos = Nothing
    Set DialogosClanes = Nothing
    Set Audio = Nothing
    Set Inventario = Nothing
    Set MainTimer = Nothing
    Set incomingData = Nothing
    Set outgoingData = Nothing
    Set Barco(0) = Nothing
    Set Barco(1) = Nothing
    UserEmbarcado = False
    
#If SeguridadAlkon Then
    Set MD5 = Nothing
#End If
    
    Call UnloadAllForms
    
    'Actualizar tip
    Config_Inicio.tip = tipf
    Call EscribirGameIni(Config_Inicio)
    
    End
End Sub
Public Function BuscarZona(ByVal x As Integer, ByVal Y As Integer) As Integer
Dim i As Integer
Dim Encontro As Boolean
Dim NewMidi As Integer
Encontro = False
For i = 1 To NumZonas
    If UserMap = Zonas(i).Mapa And x >= Zonas(i).X1 And x <= Zonas(i).X2 And Y >= Zonas(i).Y1 And Y <= Zonas(i).Y2 Then
        BuscarZona = i
        Encontro = True
        If Zonas(i).Acoplar = 0 Then Exit For
    End If
Next i
If Not Encontro And UserMap > 0 Then
    i = IIf(HayAgua(x, Y), 24, 23)
    BuscarZona = i
End If
End Function
Public Sub CheckZona()
Dim i As Integer
Dim Encontro As Boolean
Dim NewMidi As Integer
Encontro = False
For i = 1 To NumZonas
    If UserMap = Zonas(i).Mapa And UserPos.x >= Zonas(i).X1 And UserPos.x <= Zonas(i).X2 And UserPos.Y >= Zonas(i).Y1 And UserPos.Y <= Zonas(i).Y2 Then
        If ZonaActual <> i Then
            If ZonaActual > 0 Then
                If Zonas(ZonaActual).Segura <> Zonas(i).Segura Then
                    CambioSegura = True
                Else
                    CambioSegura = False
                End If
            Else
                CambioSegura = True
            End If
            ZonaActual = i
            
        End If
        Encontro = True
        If Zonas(i).Acoplar = 0 Then Exit For
    End If
Next i
If Not Encontro And UserMap > 0 Then
    i = IIf(HayAgua(UserPos.x, UserPos.Y), 24, 23)
    If ZonaActual <> i Then
        ZonaActual = i
    End If
End If
If ZonaActual > 0 Then
    If LastZona <> Zonas(ZonaActual).nombre Then
        CambioZona = 500
        If Zonas(ZonaActual).CantMusica > 0 Then
            NewMidi = Zonas(ZonaActual).Musica(RandomNumber(1, Zonas(ZonaActual).CantMusica))
            If NewMidi <> MidiCambio Then
                MidiCambio = NewMidi
                If VolumenCambio = 0 Then VolumenCambio = 100
            End If
        End If
        LastZona = Zonas(ZonaActual).nombre
    End If
End If
End Sub
Sub ClosePj()
    'Stop audio
    Dim i As Integer
    Call Audio.Sound_Stop
    frmMain.IsPlaying = PlayLoop.plNone
    
    Dim x As Integer
    Dim Y As Integer
    For x = XMinMapSize To XMaxMapSize
        For Y = YMinMapSize To YMaxMapSize
            MapData(x, Y).CharIndex = 0
            If MapData(x, Y).ObjGrh.GrhIndex = GrhFogata Then
                MapData(x, Y).Graphic(3).GrhIndex = 0
                Call Light_Destroy_ToMap(x, Y)
            End If
            MapData(x, Y).ObjGrh.GrhIndex = 0
        Next Y
    Next x
    On Local Error Resume Next
    frmMain.SendTxt.Visible = False
    frmMain.SendCMSTXT.Visible = False
    
    FrameUseMotionBlur = False
    TiempoHome = 0
    GoingHome = 0
    AngMareoMuerto = 0
    RadioMareoMuerto = 0
    BlurIntensity = 255
    ZoomLevel = 0
    'D3DDevice.SetRenderTarget pBackbuffer, DeviceStencil, 0
    
    For i = 0 To Forms.Count - 1
        If Forms(i).Name <> frmMain.Name And Forms(i).Name <> frmCrearPersonaje.Name And Forms(i).Name <> frmMensaje.Name Then
            Unload Forms(i)
        End If
    Next i
    'Show connection form
    If Not frmCrearPersonaje.Visible And Not Conectar Then
        ShowConnect
    End If
    
    'Reset global vars
    UserDescansar = False
    UserParalizado = False
    pausa = False
    UserCiego = False
    UserMeditar = False
    UserNavegando = False
    UserEmbarcado = False
    Set Barco(0) = Nothing
    Set Barco(1) = Nothing
    bRain = False
    bFogata = False
    SkillPoints = 0
    TiempoRetos = 0
    
    For i = 1 To NUMSKILLS
        UserSkills(i) = 0
    Next i

    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i
    
    frmMain.macrotrabajo.Enabled = False
    
    'Delete all kind of dialogs
    Call CleanDialogs
    
    'Reset some char variables...
    For i = 1 To LastChar
        charlist(i).invisible = False
    Next i
    
    'Unload all forms except frmMain
    Dim frm As Form
    
    For Each frm In Forms
        If frm.Name <> frmMain.Name Then
            Unload frm
        End If
    Next
    
    DoConectar
End Sub

Public Function General_Distance_Get(ByVal X1 As Integer, ByVal Y1 As Integer, X2 As Integer, Y2 As Integer) As Integer
Dim Dist As Long
Dist = Sqr((X1 - X2) ^ 2 + (Y1 - Y2) ^ 2)
General_Distance_Get = Dist
End Function
