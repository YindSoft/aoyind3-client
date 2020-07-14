Attribute VB_Name = "modEngine"
Option Explicit
Public dX As DirectX8
Public D3D As Direct3D8
Public D3DX As D3DX8
Public D3DDevice As Direct3DDevice8

Public ShadowColor As Long
Private Const FVF = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE 'D3DFVF_XYZ Or D3DFVF_NORMAL Or D3DFVF_TEX1 ' D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR
Private Const PI As Single = 3.14159275180032   'can be worked out using (4*atn(1))
Public Const ANSI_FIXED_FONT As Long = 11
Public Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
'Very percise counter 64bit system counter
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long


Public Type TLVERTEX
    x As Single
    y As Single
    z As Single
    rhw As Single
    color As Long
    tu As Single
    tv As Single
End Type
Public Type TexInfo
    x As Integer
    y As Integer
End Type
Private Type PosC
    x As Long
    y As Long
    X2 As Long
    Y2 As Long
End Type


Public SurfaceSize() As TexInfo
'The size of a FVF vertex
Private Const FVF_Size As Long = 28

Public MainFont As D3DXFont
Public MainFontDesc As IFont
Public fnt As New StdFont

Public FontCartel As D3DXFont
Public FontCartelDesc As IFont
Public fntCartel As New StdFont

Public MainFontBig As D3DXFont
Public MainFontBigDesc As IFont
Public fnt2 As New StdFont

Public pRenderTexture As Direct3DTexture8
Public pRenderSurface As Direct3DSurface8
Public pBackbuffer As Direct3DSurface8

Public Const DegreeToRadian As Single = 0.01745329251994 'Pi / 180
Public Const RadianToDegree As Single = 57.2958279087977 '180 / Pi

Private LastTexture As Integer
Public Textura As Direct3DTexture8
Private Caracteres(255) As PosC

Public Sprite As D3DXSprite
Private SpriteScaleVector As D3DVECTOR2


Public RectJuego As D3DRECT

Dim end_time As Currency
Dim timer_freq As Currency

Public Function DDRect(x, y, X1, Y1) As RECT
DDRect.Bottom = Y1
DDRect.Top = y
DDRect.Left = x
DDRect.Right = X1
End Function
Public Function IniciarD3D() As Boolean
On Error Resume Next
Set dX = New DirectX8
    If Err Then
        MessageBox "No se puede iniciar DirectX. Por favor asegurese de tener la ultima version correctamente instalada."
        Exit Function
    End If

Set D3D = dX.Direct3DCreate()
Set D3DX = New D3DX8

If Not IniciarDevice(D3DCREATE_PUREDEVICE) Then
    If Not IniciarDevice(D3DCREATE_HARDWARE_VERTEXPROCESSING) Then
        If Not IniciarDevice(D3DCREATE_MIXED_VERTEXPROCESSING) Then
            If Not IniciarDevice(D3DCREATE_SOFTWARE_VERTEXPROCESSING) Then
                MessageBox "No se pudo iniciar el D3DDevice. Saliendo...", vbCritical
                LiberarObjetosDX
                End
            End If
        End If
    End If
End If

Call SetDevice(D3DDevice)




Set Sprite = D3DX.CreateSprite(D3DDevice)
     
'Set the scaling to default aspect ratio
SpriteScaleVector.x = 1
SpriteScaleVector.y = 1

Call setup_ambient

IluRGB.R = 255
IluRGB.G = 255
IluRGB.B = 255

Iluminacion = D3DColorRGBA(IluRGB.R, IluRGB.G, IluRGB.B, 255)
ColorTecho = Iluminacion

bAlpha = 255

Set pRenderTexture = D3DX.CreateTexture(D3DDevice, 480, 80, 1, D3DUSAGE_RENDERTARGET, D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT)
Set pRenderSurface = pRenderTexture.GetSurfaceLevel(0)
Set pBackbuffer = D3DDevice.GetRenderTarget

IniciarD3D = True
End Function
Public Sub SetDevice(D3DD As Direct3DDevice8)
With D3DD
    .SetVertexShader FVF

    'Set the render states
    .SetRenderState D3DRS_LIGHTING, False
    .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    .SetRenderState D3DRS_ALPHABLENDENABLE, True
    .SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
    .SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
    .SetRenderState D3DRS_ZENABLE, True
    .SetRenderState D3DRS_ZWRITEENABLE, True
    .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE

    'Particle engine settings
    .SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
    .SetRenderState D3DRS_POINTSCALE_ENABLE, 0
    
    'Set the texture stage stats (filters)
    .SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_POINT
    .SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_POINT
End With
End Sub
Public Sub CargarFont()
Dim i As Integer
Open (App.path & "\Init\Font.ind") For Binary As #1
    For i = 1 To 255
        Get #1, , Caracteres(i)
    Next i
Close #1

Dim hFont As Long

fntCartel.Name = "Augusta"
fntCartel.Size = 14
fntCartel.bold = False
Set FontCartelDesc = fntCartel


fnt.Name = "Augusta"
fnt.Size = 48
fnt.bold = False
Set MainFontDesc = fnt

fnt2.Name = "Augusta"
fnt2.Size = 72
fnt2.bold = False
Set MainFontBigDesc = fnt2
'hFont = GetStockObject(ANSI_FIXED_FONT)
    
Set MainFont = D3DX.CreateFont(D3DDevice, MainFontDesc.hFont)
Set MainFontBig = D3DX.CreateFont(D3DDevice, MainFontBigDesc.hFont)
Set FontCartel = D3DX.CreateFont(D3DDevice, FontCartelDesc.hFont)
End Sub
Public Sub DrawFont(Texto As String, ByVal x As Long, ByVal y As Long, ByVal color As Long, Optional Centrado As Boolean = False)
Dim i As Integer
Dim SumaX As Integer
Dim SumaL As Integer
Dim CharC As Byte
If Centrado Then
    For i = 1 To Len(Texto)
        CharC = Asc(mid$(Texto, i, 1))
        SumaL = SumaL + Caracteres(CharC).X2 - 2
    Next i
    SumaL = SumaL / 2
End If
For i = 1 To Len(Texto)
    CharC = Asc(mid$(Texto, i, 1))
    'Call Engine_Render_D3DXSprite(X - SumaL + SumaX, Y, Caracteres(CharC).X2 + 2, Caracteres(CharC).Y2 + 2, Caracteres(CharC).X + 1, Caracteres(CharC).Y + 1, Color, 14324, 0)
    Call Engine_Render_Rectangle(x - SumaL + SumaX, y, Caracteres(CharC).X2 + 2, Caracteres(CharC).Y2 + 2, Caracteres(CharC).x + 1, Caracteres(CharC).y + 1, Caracteres(CharC).X2 + 2, Caracteres(CharC).Y2 + 2, , , , 14324, color, color, color, color)
    
    SumaX = SumaX + Caracteres(CharC).X2 - 2
Next i
End Sub
Public Function IniciarDevice(D3DCREATEFLAGS As CONST_D3DCREATEFLAGS) As Boolean
On Error GoTo ErrOut

Dim DispMode As D3DDISPLAYMODE
Dim D3DWindow As D3DPRESENT_PARAMETERS
UseMotionBlur = 1

D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode

D3DWindow.Windowed = 1
D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY
D3DWindow.BackBufferFormat = DispMode.Format
D3DWindow.EnableAutoDepthStencil = 0
D3DWindow.AutoDepthStencilFormat = D3DFMT_A8R8G8B8

'If UseMotionBlur Then
'    D3DWindow.EnableAutoDepthStencil = 1
'    D3DWindow.AutoDepthStencilFormat = D3DFMT_D16
'End If
frmMain.SetRender (True)

RectJuego.X1 = 0
RectJuego.Y1 = 0
RectJuego.X2 = 736
RectJuego.Y2 = 544


If Not D3DDevice Is Nothing Then Set D3DDevice = Nothing
Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.pRender.hwnd, D3DCREATEFLAGS, D3DWindow)

    
    If UseMotionBlur Then
        'Set DeviceBuffer = D3DDevice.GetRenderTarget
        'Set DeviceStencil = D3DDevice.GetDepthStencilSurface
        'Set BlurStencil = D3DDevice.CreateDepthStencilSurface(800, 600, D3DFMT_D16, D3DMULTISAMPLE_NONE)
        Set BlurTexture = D3DX.CreateTexture(D3DDevice, 1024, 1024, 1, D3DUSAGE_RENDERTARGET, D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT)
        Set BlurSurf = BlurTexture.GetSurfaceLevel(0)
       
        Dim T As Long
       
        'Create the motion-blur vertex array
        For T = 0 To 3
            BlurTA(T).color = D3DColorXRGB(255, 255, 255)
            BlurTA(T).rhw = 1
        Next T
        BlurTA(1).x = ScreenWidth
        BlurTA(2).y = ScreenHeight
        BlurTA(3).x = ScreenWidth
        BlurTA(3).y = ScreenHeight
       
    End If
   
    'Set the blur to off
    BlurIntensity = 255

IniciarDevice = True
Exit Function

ErrHandler:
MessageBox "Su placa de video no es combatible. Este al tanto en la página web para parches que puedan solucionar este incomveniente.", vbCritical
IniciarDevice = False

Exit Function

ErrOut:

    'Destroy the D3DDevice so it can be remade
    Set D3DDevice = Nothing

    'Return a failure
    IniciarDevice = False

End Function

Public Sub LiberarObjetosDX()
Err.Clear
On Error GoTo fin:

Set D3DDevice = Nothing
Set D3D = Nothing
Set D3DX = Nothing
Set dX = Nothing
Exit Sub
fin: MsgBox "Error producido en Public Sub LiberarObjetosDX()"
End Sub

Public Sub Engine_ReadyTexture(ByVal TextureNum As Integer)
    'Set the texture
    If TextureNum > 0 Then
        If LastTexture <> TextureNum Then
            Set Textura = SurfaceDB.Surface(TextureNum)
            D3DDevice.SetTexture 0, Textura
            LastTexture = TextureNum
        End If
    End If
    LastTexture = TextureNum
End Sub

Public Sub Engine_Render_D3DXSprite(ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single, ByVal srcX As Single, ByVal srcY As Single, ByVal Light As Long, ByVal TextureNum As Long, ByVal Degrees As Single)
Dim SrcRect As RECT
Dim v2 As D3DVECTOR2
Dim v3 As D3DVECTOR2

    
    'Ready the texture
    Engine_ReadyTexture TextureNum
    
    'Create the source rectangle
    With SrcRect
        .Left = srcX
        .Top = srcY
        .Right = .Left + Width
        .Bottom = .Top + Height
    End With
    
    'Create the rotation point
    If Degrees Then
        Degrees = ((Degrees + 180) * DegreeToRadian)
        If Degrees > 360 Then Degrees = Degrees - 360
        With v2
            .x = (Width * 0.5)
            .y = (Height * 0.5)
        End With
    End If
    
    'Set the translation (location on the screen)
    v3.x = x - 256
    v3.y = y - 256

    'Draw the sprite
    If TextureNum > 0 Then
        Sprite.Draw Textura, SrcRect, SpriteScaleVector, v2, Degrees, v3, Light
    Else
        'Sprite.Draw Nothing, SrcRect, SpriteScaleVector, v2, 0, v3, Light
    End If
    
End Sub

Public Sub Engine_Render_D3DXTexture(ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single, ByVal srcX As Single, ByVal srcY As Single, ByVal Light As Long, ByVal Texture As Direct3DTexture8, ByVal Degrees As Single)
Dim SrcRect As RECT
Dim v2 As D3DVECTOR2
Dim v3 As D3DVECTOR2

    
    'Ready the texture
    'D3DDevice.SetTexture 0, Texture
    LastTexture = 0
    
    'Create the source rectangle
    With SrcRect
        .Left = srcX
        .Top = srcY
        .Right = .Left + Width
        .Bottom = .Top + Height
    End With
    
    'Create the rotation point
    If Degrees Then
        Degrees = ((Degrees + 180) * DegreeToRadian)
        If Degrees > 360 Then Degrees = Degrees - 360
        With v2
            .x = (Width * 0.5)
            .y = (Height * 0.5)
        End With
    End If
    
    'Set the translation (location on the screen)
    v3.x = x - 256
    v3.y = y - 256

    'Draw the sprite
    Sprite.Draw Texture, SrcRect, SpriteScaleVector, v2, Degrees, v3, Light
    
End Sub

Sub Engine_Render_Rectangle(ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single, ByVal srcX As Single, ByVal srcY As Single, ByVal SrcWidth As Single, ByVal SrcHeight As Single, Optional ByVal SrcBitmapWidth As Long = -1, Optional ByVal SrcBitmapHeight As Long = -1, Optional ByVal Degrees As Single = 0, Optional ByVal TextureNum As Long, Optional ByVal Color0 As Long = -1, Optional ByVal Color1 As Long = -1, Optional ByVal Color2 As Long = -1, Optional ByVal Color3 As Long = -1, Optional ByVal Shadow As Byte = 0, Optional ByVal InBoundsCheck As Boolean = True)
'************************************************************
'Render a square/rectangle based on the specified values then rotate it if needed
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Render_Rectangle
'************************************************************
Dim VertexArray(0 To 3) As TLVERTEX
Dim RadAngle As Single 'The angle in Radians
Dim CenterX As Single
Dim CenterY As Single
Dim Index As Integer
Dim NewX As Single
Dim NewY As Single
Dim SinRad As Single
Dim CosRad As Single
Dim ShadowAdd As Single
Dim l As Single

    'Perform in-bounds check if needed
    If InBoundsCheck Then
        If x - 256 + SrcWidth <= 0 Then Exit Sub
        If y - 256 + SrcHeight <= 0 Then Exit Sub
        If x - 256 >= frmMain.pRender.Width Then Exit Sub
        If y - 256 >= frmMain.pRender.Height Then Exit Sub
    End If

    'Ready the texture
    Engine_ReadyTexture TextureNum

    'Set the bitmap dimensions if needed
    If SrcBitmapWidth = -1 Then SrcBitmapWidth = SurfaceSize(TextureNum).x
    If SrcBitmapHeight = -1 Then SrcBitmapHeight = SurfaceSize(TextureNum).y
    
    'Set the RHWs (must always be 1)
    VertexArray(0).rhw = 1
    VertexArray(1).rhw = 1
    VertexArray(2).rhw = 1
    VertexArray(3).rhw = 1
    
    'Apply the colors
    VertexArray(0).color = Color0
    VertexArray(1).color = Color1
    VertexArray(2).color = Color2
    VertexArray(3).color = Color3

    If Shadow Then

        'To make things easy, we just do a completely separate calculation the top two points
        ' with an uncropped tU / tV algorithm
        VertexArray(0).x = x - 256 + (Width * 0.5)
        VertexArray(0).y = y - 256 - (Height * 0.5)
        VertexArray(0).tu = (srcX / SrcBitmapWidth)
        VertexArray(0).tv = (srcY / SrcBitmapHeight)
        
        VertexArray(1).x = VertexArray(0).x + Width
        VertexArray(1).tu = ((srcX + Width) / SrcBitmapWidth)

        VertexArray(2).x = x - 256
        VertexArray(2).tu = (srcX / SrcBitmapWidth)

        VertexArray(3).x = x - 256 + Width
        VertexArray(3).tu = (srcX + SrcWidth + ShadowAdd) / SrcBitmapWidth

    Else
        
        '------------------------------------------------------------------------------------------------------
        '------------------------------------------------------------------------------------------------------
        'If the image is partially outside of the screen, it is trimmed so only that which is in the screen is drawn
        'This provides for quite a decent FPS boost if you have lots of tiles that stretch outside of the view area
        'Important: Something about this doesn't seem to be functioning correctly. It is supposed to crop down the
        'image and only draw that which is going to be in the screen, but it doesn't work right and I have no
        'idea why. Uncomment the lines to see what happens. I have given up on this since the FPS boost really isn't
        'significant for me to put any more work into it, but if someone could fix it, it would definitely be
        'added back into the engine.
        '------------------------------------------------------------------------------------------------------
        '------------------------------------------------------------------------------------------------------
        'If X < 0 Then
        '    SrcX = SrcX - X
        '    SrcWidth = SrcWidth + X
        '    Width = Width + X
        '    X = 0
        'End If
        'If Y < 0 Then
        '    SrcY = SrcY - Y
        '    SrcHeight = SrcHeight + Y
        '    Height = Height + Y
        '    Y = 0
        'End If
        'If X + Width > ScreenWidth Then
        '    L = X + Width - ScreenWidth
        '    Width = Width - L
        '    SrcWidth = SrcWidth - L
        'End If
        'If Y + Height > ScreenHeight Then
        '    L = Y + Height - ScreenHeight
        '    Height = Height - L
        '    SrcHeight = SrcHeight - L
        'End If
        '------------------------------------------------------------------------------------------------------
        '------------------------------------------------------------------------------------------------------
        
        'If we are NOT using shadows, then we add +1 to the width/height (trust me, just do it)
        ShadowAdd = 1

        'Find the left side of the rectangle
        VertexArray(0).x = x - 256
        If SrcBitmapWidth = 0 Then Exit Sub
        VertexArray(0).tu = (srcX / SrcBitmapWidth)

        'Find the top side of the rectangle
        VertexArray(0).y = y - 256
        VertexArray(0).tv = (srcY / SrcBitmapHeight)
    
        'Find the right side of the rectangle
        VertexArray(1).x = x - 256 + Width
        VertexArray(1).tu = (srcX + SrcWidth + ShadowAdd) / SrcBitmapWidth

        'These values will only equal each other when not a shadow
        VertexArray(2).x = VertexArray(0).x
        VertexArray(3).x = VertexArray(1).x
    End If
    
    'Find the bottom of the rectangle
    VertexArray(2).y = y - 256 + Height
    VertexArray(2).tv = (srcY + SrcHeight + ShadowAdd) / SrcBitmapHeight

    'Because this is a perfect rectangle, all of the values below will equal one of the values we already got
    VertexArray(1).y = VertexArray(0).y
    VertexArray(1).tv = VertexArray(0).tv
    VertexArray(2).tu = VertexArray(0).tu
    VertexArray(3).y = VertexArray(2).y
    VertexArray(3).tu = VertexArray(1).tu
    VertexArray(3).tv = VertexArray(2).tv
    
    'Check if a rotation is required
    If Degrees Mod 360 <> 0 Then

        'Converts the angle to rotate by into radians
        RadAngle = Degrees * DegreeToRadian

        'Set the CenterX and CenterY values
        CenterX = x - 256 + (Width * 0.5)
        CenterY = y - 256 + (Height * 0.5)

        'Pre-calculate the cosine and sine of the radiant
        SinRad = Sin(RadAngle)
        CosRad = Cos(RadAngle)

        'Loops through the passed vertex buffer
        For Index = 0 To 3

            'Calculates the new X and Y co-ordinates of the vertices for the given angle around the center co-ordinates
            NewX = CenterX + (VertexArray(Index).x - CenterX) * CosRad - (VertexArray(Index).y - CenterY) * SinRad
            NewY = CenterY + (VertexArray(Index).y - CenterY) * CosRad + (VertexArray(Index).x - CenterX) * SinRad

            'Applies the new co-ordinates to the buffer
            VertexArray(Index).x = NewX
            VertexArray(Index).y = NewY

        Next Index

    End If

    'Render the texture to the device
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), FVF_Size

End Sub
