Attribute VB_Name = "modPresentacion"
Option Explicit
'Presentacion
Public GTCPres As Long, GTCInicial As Long, GTCChars As Long
Public Conectar As Boolean
Public MostrarEntrar As Long
Private MouseOn As Integer
Public logUser As String
Public logPass As String
Dim CallDoConectar As Boolean
Dim IntroChars() As Char
Dim CantPjs As Integer
Dim MousePj As Integer
Dim AperturaX As Integer
Dim AperturaY As Integer
Dim AperturaPj As Integer
Dim AperturaTick As Long


Public UseMotionBlur As Byte    'If motion blur is enabled or not
Public BlurIntensity As Single
Public BlurTexture As Direct3DTexture8
Public BlurSurf As Direct3DSurface8
Public BlurStencil As Direct3DSurface8
Public DeviceStencil As Direct3DSurface8
Public DeviceBuffer As Direct3DSurface8
Public BlurTA(0 To 3) As TLVERTEX
 
'Zoom level - 0 = No Zoom, > 0 = Zoomed
Public ZoomLevel As Single
Public Const MaxZoomLevel As Single = 0.183
 
Public Const ScreenWidth As Long = 1024 'anchonho
Public Const ScreenHeight As Long = 768 'altongo d su render





Public Sub MouseAction(x As Single, y As Single, Button As Integer)
If GTCPres > 7000 Then 'Me aseguro que este todo cargado
    If MostrarEntrar > 0 Then 'Si esta abierto el cuadro entrar
        If x >= 145 + 112 And x <= 660 + 112 And y >= 345 + 168 And y <= 555 + 168 Then ' Cuadro entrar
            If x >= 500 + 112 And x <= 545 + 112 And y >= 432 + 168 And y <= 484 + 168 Then
                If Button = 1 Then
                    ClickAbrirCuenta
                    Call Audio.Sound_Play(SND_CLICKNEW)
                Else
                    If MouseOn <> 2 Then
                        MouseOn = 2
                        Call Audio.Sound_Play(SND_MOUSEOVER)
                    End If
                End If
            Else
                MouseOn = 0
            End If
        Else 'Si el mouse esta fuera del cuadro entrar
            If Button = 1 Then
                Call Audio.Sound_Play(SND_CLICKOFF)
            End If
            MouseOn = 0
        End If
    ElseIf MostrarEntrar = 0 Then  'Si no esta abierto el cuadro entrar
        If AperturaPj > 0 And Button = 1 Then
            CloseSock
            AperturaPj = -AperturaPj
            AperturaTick = (GetTickCount() And &H7FFFFFFF)
        ElseIf x >= 355 + 112 And x <= 450 + 112 And y >= 130 And y <= 160 Then 'Boton Entrar
            If Button = 1 Then
                MostrarEntrar = GTCPres
                Call Audio.Sound_Play(SND_CLICKNEW)
                Call Audio.Sound_Play(SND_CADENAS)
            Else
                If MouseOn <> 1 Then
                    MouseOn = 1
                    Call Audio.Sound_Play(SND_MOUSEOVER)
                End If
            End If
        ElseIf x >= 15 + 112 And x <= 105 + 112 And y >= 50 And y <= 75 Then 'Boton crear
            If Button = 1 Then
                frmNavegador.TIPO = Crear
                frmNavegador.Show vbModal
                Call Audio.Sound_Play(SND_CLICKNEW)
            Else
                If MouseOn <> 1 Then
                    MouseOn = 1
                    Call Audio.Sound_Play(SND_MOUSEOVER)
                End If
            End If
        ElseIf x >= 121 + 112 And x <= 229 + 112 And y >= 50 And y <= 75 Then 'Boton recuperar
            If Button = 1 Then
                frmNavegador.TIPO = Recuperar
                frmNavegador.Show vbModal
                Call Audio.Sound_Play(SND_CLICKNEW)
            Else
                If MouseOn <> 1 Then
                    MouseOn = 1
                    Call Audio.Sound_Play(SND_MOUSEOVER)
                End If
            End If
        ElseIf x >= 576 + 112 And x <= 668 + 112 And y >= 50 And y <= 75 Then 'Boton borrar
            If Button = 1 Then
                frmNavegador.TIPO = Borrar
                frmNavegador.Show vbModal
                Call Audio.Sound_Play(SND_CLICKNEW)
            Else
                If MouseOn <> 1 Then
                    MouseOn = 1
                    Call Audio.Sound_Play(SND_MOUSEOVER)
                End If
            End If
        ElseIf x >= 693 + 112 And x <= 783 + 112 And y >= 50 And y <= 75 Then 'Boton salir
            If Button = 1 Then
                prgRun = False
                Call Audio.Sound_Play(SND_CLICKNEW)
            Else
                If MouseOn <> 1 Then
                    MouseOn = 1
                    Call Audio.Sound_Play(SND_MOUSEOVER)
                End If
            End If
        ElseIf x >= 105 + 112 And x <= 200 + 112 And y >= 130 And y <= 160 Then
            If Button = 1 Then

            Else
                If MouseOn <> 1 Then
                    MouseOn = 1
                    Call Audio.Sound_Play(SND_MOUSEOVER)
                End If
            End If
        ElseIf CantPjs > 0 Then
            Dim i As Integer
            Dim Angulo As Single
        
            MousePj = 0
            For i = 1 To CantPjs
                Angulo = (-40 * CantPjs + i * 80 - 48) / 180 - 1.57
                If Abs(x - (512 + Cos(Angulo) * 320 + 16)) < 32 And Abs(y - (450 + Sin(Angulo) * 160)) < 54 Then
                    MousePj = i
                End If
            Next i
            
            'If x >= 400 - CantPjs * 40 And x <= 400 + CantPjs * 40 And y >= 250 And y <= 350 Then
            '    MousePj = (x - 400 + CantPjs * 40 - 48) / 80 + 1
            'Else
            '    MousePj = 0
            'End If
            If MousePj > 0 And Button = 1 Then
                If IntroChars(MousePj).priv = 9 Then
                    frmCrearPersonaje.Show , frmMain
                    Call Audio.Sound_Play(SND_CLICKNEW)
                Else
                    UserName = IntroChars(MousePj).nombre
                    AperturaPj = MousePj
                    AperturaTick = (GetTickCount() And &H7FFFFFFF)
                    EstadoLogin = Normal
                    'Login
                    
                    iServer = 0
                    iCliente = 0
                    DummyCode = StrConv(StrReverse("conectar") & "CuEnTa", vbFromUnicode)
                    DoEvents
                    If Not ClientSetup.WinSock Then
                        frmMain.Client.CloseSck
                        frmMain.Client.Connect IpServidor, 7222
                    Else
                        frmMain.WSock.Close
                        frmMain.WSock.Connect IpServidor, 7222
                    End If
                End If
            End If
        Else
            MouseOn = 0
        End If
    End If
End If
End Sub
Public Sub ClickAbrirCuenta()
logUser = frmMain.tUser.Text
logPass = frmMain.tPass.Text

DoConectar
End Sub
Public Sub DoConectar()
CloseSock
DoEvents
'update user info
UserName = logUser
Dim aux As String
aux = logPass
UserPassword = MD5(aux)
UserAccount = logUser
iCliente = 0
iServer = 0
DummyCode = StrConv(StrReverse("conectar") & "CuEnTa", vbFromUnicode)
If CheckUserData(False) = True Then
    EstadoLogin = Cuentas
    
    If Not ClientSetup.WinSock Then
        frmMain.Client.Connect IpServidor, 7222
    Else
        frmMain.WSock.Connect IpServidor, 7222
    End If
End If
End Sub
Sub RenderConectar()
Static Ang As Single
Dim color As Long
GTCPres = Abs((GetTickCount() And &H7FFFFFFF) - GTCInicial)

If GTCPres < 4000 Then
    color = D3DColorRGBA(255, 255, 255, CalcAlpha(GTCPres, 0, 220, 15))
    Call Engine_Render_Rectangle(255 + 412, 255 + 284, 200, 200, 0, 0, 400, 400, , , 0, 14705, color, color, color, color)
End If

Dim a As Single
Dim T As Single
Dim T2 As Single
Dim Mueve As Single
Dim tmpColor As Long
Dim i As Integer
Dim Angulo As Single
Dim Separacion As Single
Dim x As Single
Dim y As Single
Static AngSel As Single

a = 20
T = (GTCPres - 4000) / 1000

If GTCPres >= 4000 Then
    If MostrarEntrar > 0 Then
        tmpColor = 255 - CalcAlpha(GTCPres, MostrarEntrar, 170, 5)
    ElseIf MostrarEntrar < 0 Then
        tmpColor = 85 + CalcAlpha(GTCPres, -MostrarEntrar, 170, 5)
    Else
        tmpColor = 255
    End If
    Call Engine_Render_D3DXSprite(255, 255, 1024, 768, 0, 0, D3DColorRGBA(tmpColor, tmpColor, tmpColor, CalcAlpha(GTCPres, 4000, 255, 15)), 14704, 0)
    
    tmpColor = D3DColorRGBA(255, 255, 255, 220 - CalcAlpha(GTCPres, 4000, 220, 15))
    Call Engine_Render_Rectangle(255 + 412, 255 + 284 - a * (T ^ 2), 200, 200, 0, 0, 400, 400, , , 0, 14705, tmpColor, tmpColor, tmpColor, tmpColor)

If CantPjs > 0 Then
    For i = 1 To CantPjs
        T2 = Abs((GetTickCount() And &H7FFFFFFF) - GTCChars) / 1000
        Separacion = 120 * T2 - 40 * (T2 ^ 2)
        If T2 > 1 Then Separacion = 80
        If IntroChars(i).Alpha < 160 Then IntroChars(i).Alpha = IntroChars(i).Alpha + 5
        If (AperturaPj <= 0 And MousePj = i) Or AperturaPj = i Then
            IntroChars(i).Alpha = 255
        Else
            IntroChars(i).Alpha = 160
        End If
        'Call DrawFont(CStr(Separacion), 323, 423, D3DColorRGBA(255, 255, 255, 160))
        'Call CharRender(IntroChars(i), -1, 255 + 400 - Separacion / 2 * CantPjs + i * Separacion - 48 * Separacion / 80, 255 + 300)
        
        Angulo = (-Separacion / 2 * CantPjs + i * Separacion - 48 * Separacion / 80) / 180 - 1.57
        
        T2 = Abs((GetTickCount() And &H7FFFFFFF) - AperturaTick) / 1000
        If AperturaPj > 0 And AperturaX < 660 Then
            AperturaX = 320 + (T2 ^ 2) * 550
            AperturaY = 160 + (T2 ^ 2) * 412.5
            If AperturaX >= 660 Then
                AperturaX = 660
                AperturaY = 441
            End If
        ElseIf AperturaPj < 0 And AperturaX > 320 Then
            AperturaX = 660 - (T2 ^ 2) * 550
            AperturaY = 441 - (T2 ^ 2) * 412.5
            If AperturaX <= 320 Then
                AperturaX = 320
                AperturaY = 160
            End If
        End If
        
        If i = AperturaPj Or i = -AperturaPj Then
            x = 255 + 497 + Cos(Angulo) * (497 - AperturaX / 2)
            y = 255 + 412 + Sin(Angulo) * (147 - AperturaY / 3) - (AperturaY - 110) / 6
        Else
            x = 255 + 512 + Cos(Angulo) * AperturaX * 1.5
            y = 255 + 450 + Sin(Angulo) * AperturaY * 1.5
        End If
        
        If IntroChars(i).logged Then
            IntroChars(i).Alpha = 70
        End If
        
        Call CharRender(IntroChars(i), -1, x, y)
        a = a + 1
        If (AperturaPj <= 0 And MousePj = i) Or AperturaPj = i Then
            D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
            D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
            AngSel = AngSel + 0.05
                Call Engine_Render_Rectangle(x - 47, y - 52, 128, 128, 224, 0, 128, 128, , , AngSel, 14332)
                             
            D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
            D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        End If
    Next i
End If


    Mueve = (T * 20) Mod 512
    
    tmpColor = D3DColorRGBA(255, 255, 255, CalcAlpha(GTCPres, 4000, 150, 15))
    
    
    Call Engine_Render_D3DXSprite(255, 255, 512 - Mueve, 512, Mueve, 0, tmpColor, 14706, 0)
    Call Engine_Render_D3DXSprite(255, 767, 512 - Mueve, 256, Mueve, 0, tmpColor, 14706, 0)
    
    Call Engine_Render_D3DXSprite(767 - Mueve, 255, 512, 512, 0, 0, tmpColor, 14706, 0)
    Call Engine_Render_D3DXSprite(767 - Mueve, 767, 512, 256, 0, 0, tmpColor, 14706, 0)
    
    Call Engine_Render_D3DXSprite(1279 - Mueve, 255, Mueve, 512, 0, 0, tmpColor, 14706, 0)
    Call Engine_Render_D3DXSprite(1279 - Mueve, 767, Mueve, 256, 0, 0, tmpColor, 14706, 0)
        

    If MostrarEntrar > 0 Then
        T2 = (GTCPres - MostrarEntrar) / 1000
        If T2 < 1 Then
            Call Engine_Render_D3DXSprite(255, 1023 - 388.5 * T2 + 259 / 2 * (T2 ^ 2), 1024, 259, 0, 177, D3DColorRGBA(255, 255, 255, CalcAlpha(GTCPres, MostrarEntrar, 200, 4)), 14703, 0)
        Else
            Call Engine_Render_D3DXSprite(255, 1023 - 259, 1024, 259, 0, 177, D3DColorRGBA(255, 255, 255, 200), 14703, 0)
            If frmMain.tUser.Visible = False Then
                frmMain.tUser.Text = ""
                frmMain.tPass.Text = ""
                frmMain.tUser.Visible = True
                frmMain.tPass.Visible = True
                frmMain.tUser.SetFocus
            End If
        End If
    ElseIf MostrarEntrar < 0 Then
        T2 = (GTCPres + MostrarEntrar) / 1000
        If T2 < 1 Then
            Call Engine_Render_D3DXSprite(255, 1023 - 259 + 388.5 * T2 - 259 / 2 * (T2 ^ 2), 800, 259, 0, 177, D3DColorRGBA(255, 255, 255, 200 - CalcAlpha(GTCPres, -MostrarEntrar, 200, 4)), 14703, 0)
        Else
            MostrarEntrar = 0
            If CallDoConectar Then
                CallDoConectar = False
                MostrarEntrar = -GTCPres
                frmMain.tUser.Visible = False
                frmMain.tPass.Visible = False
                Call Audio.Sound_Play(SND_CADENAS)
            End If
        End If
    End If
    
    
    If T <= 4 Then
        Call Engine_Render_D3DXSprite(255, 255 - 177 + Int(88.5 * T - 22.125 / 2 * (T ^ 2)), 1024, 177, 0, 0, D3DColorRGBA(255, 255, 255, CalcAlpha(GTCPres, 4000, 255, 15)), 14703, 0)
        Call Engine_Render_D3DXSprite(255, 1023 - Int(23 * T - 5.75 / 2 * (T ^ 2)), 1024, 47, 0, 436, D3DColorRGBA(255, 255, 255, CalcAlpha(GTCPres, 4000, 255, 15)), 14703, 0)
    Else
        Call Engine_Render_D3DXSprite(255, 255, 1024, 177, 0, 0, D3DColorRGBA(255, 255, 255, CalcAlpha(GTCPres, 4000, 255, 15)), 14703, 0)
        Call Engine_Render_D3DXSprite(255, 1023 - 46, 1024, 47, 0, 436, D3DColorRGBA(255, 255, 255, CalcAlpha(GTCPres, 4000, 255, 15)), 14703, 0)
    End If
    
    'If T Mod 2 = 0 Then If Not MP3P.IsItPlaying Then Call EscucharMp3(10)
End If

Ang = Ang + 0.001

'Call TestPart.ReLocate(Cos(Ang * 3) * 40 + 300, Sin(Ang * 2) * 40 + 300)

'TestPart.Update
'TestPart.Render

'Call Engine_Render_D3DXTexture(255 + 209, 255 + 200, 103, 132, 0, 0, D3DColorRGBA(255, 255, 255, CalcAlpha(GTCPres, 0, 220, 15)), ImgBruma, 0)

'Particle_Group_Render MapData(150, 800, 1).particle_group_index, 400, 400

End Sub

Sub ShowConnect()
Call Audio.MusicMP3Play("10.mp3")
frmMain.SetRender (True)
GTCInicial = (GetTickCount() And &H7FFFFFFF) - 10000
GTCPres = (GetTickCount() And &H7FFFFFFF)
MouseOn = 0
MostrarEntrar = 0
Conectar = True
CantPjs = 0
ReDim IntroChars(0)
frmMain.tUser.Visible = False
frmMain.tPass.Visible = False
End Sub

Public Sub HandleOpenAccount()

    If incomingData.Length < 3 Then
        NotEnoughData = True
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    
    Dim i As Integer
    Dim Arma As Integer, Escudo As Integer, Casco As Integer
    CantPjs = incomingData.ReadInteger() + 1
    If CantPjs > 0 Then
    ReDim IntroChars(1 To CantPjs)
    
    For i = 1 To CantPjs - 1
        With IntroChars(i)
            .ACTIVE = 1
        
            .nombre = incomingData.ReadASCIIString()
            .iHead = incomingData.ReadInteger()
            .iBody = incomingData.ReadInteger()
            
            Arma = incomingData.ReadInteger()
            Escudo = incomingData.ReadInteger()
            Casco = incomingData.ReadInteger()
            
            If Arma = 0 Then Arma = 2
            If Escudo = 0 Then Escudo = 2
            If Casco = 0 Then Casco = 2
            
            .Head = HeadData(.iHead)
            .Body = BodyData(.iBody)
            .Arma = WeaponAnimData(Arma)
            .Escudo = ShieldAnimData(Escudo)
            .Casco = CascoAnimData(Casco)
            
            If .iBody = FRAGATA_FANTASMAL Then
                .Head = HeadData(2)
            End If
            
            .Heading = south
            
            .Alpha = 0

            .logged = incomingData.ReadBoolean()
            .Criminal = incomingData.ReadBoolean()
            .muerto = .iHead = CASPER_HEAD Or .iHead = CASPER_HEAD_CRIMI Or .iBody = FRAGATA_FANTASMAL
        End With
    Next i
    If CantPjs <= 10 Then 'Si tiene 10 pjs no le deja crear mas
        With IntroChars(CantPjs)
            .ACTIVE = 1
            .nombre = "CREAR PJ"
            .iHead = 10
            .iBody = 21
            .priv = 9
            
            .Head = HeadData(.iHead)
            .Body = BodyData(.iBody)
            .Arma = WeaponAnimData(2)
            .Escudo = ShieldAnimData(2)
            .Casco = CascoAnimData(2)
            
            .Heading = south
            
            .Alpha = 0
        End With
    Else
        CantPjs = CantPjs - 1
    End If
    End If
    MousePj = 0
    MostrarEntrar = -GTCPres
    frmMain.tUser.Visible = False
    frmMain.tPass.Visible = False
    GTCChars = (GetTickCount() And &H7FFFFFFF)
    
    AperturaX = 220
    AperturaY = 110
    AperturaPj = 0
    AperturaTick = 0

End Sub
