VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParticulas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Particles() As clsParticle
Private vertsPoints() As typeTRANSLITVERTEX

Private mCol As Collection

Private sngX As Single, sngY As Single
Private OffX As Single, OffY As Single

Public ParticleCounts As Long

Private lngFloat0 As Long
Private lngFloat1 As Long
Private lngFloatSize As Long

Private lngPreviousFrame As Long
Private myTexture As Integer
Private Iniciada As Boolean
Private Iniciadas As Long
Private Contador As Single
Public R As Single
Public G As Single
Public B As Single
Public Matar As Boolean
Public Desa As Single
Public Efecto As Byte
Public Angulo As Single
Public Property Let Texture(ByVal vData As Integer)
    myTexture = vData
End Property

Public Property Get Texture() As Integer
    Texture = myTexture
End Property

Public Sub Begin()
    '//We initialize our stuff here
    Dim i As Long
    
    lngFloat0 = FtoDW(0)
    lngFloat1 = FtoDW(1)
    lngFloatSize = FtoDW(16) '//Size of our particles..
    
    '//Redim our particles to the particlecount
    ReDim Particles(0 To ParticleCounts)
    
    '//Redim vertices to the particle count
    '//Point sprites, so 1 per particle
    ReDim vertsPoints(0 To ParticleCounts)
    
    Desa = 1
    '//Now generate all particles
    Iniciadas = 0
    For i = 0 To Iniciadas
        Iniciar (i)
    Next i
    
    '//Set initial time
    lngPreviousFrame = (GetTickCount() And &H7FFFFFFF)
    Iniciada = True
End Sub
Sub Iniciar(i As Long)
Set Particles(i) = New clsParticle
vertsPoints(i).rhw = 1
Call Reset(i)
End Sub

Public Sub Reset(i As Long)
    Dim x As Single, y As Single
    
    x = sngX '+ (Rnd * 10)
    y = sngY
    
    '//This is were we will reset individual particles.
    With Particles(i)
        Call .ResetIt(x, y, 0, 0, 0, 0, 2)
        
        If Efecto = 1 Then
        
            Call .ResetColor(R + Rnd * 1, G, B, 0.6 + (0.2 * Rnd), 0.08 + Rnd * 0.04)
        ElseIf Efecto = 2 Then
            Call .ResetColor(R, G, B, 1, 0)
        End If
    End With
End Sub

Public Sub Update()
    Dim i As Long
    Dim sngElapsedTime As Single
    
    If Not Iniciada Then Begin
    
    '//We calculate the time difference here
    sngElapsedTime = ((GetTickCount() And &H7FFFFFFF) - lngPreviousFrame) / 100
    If sngElapsedTime = 0 Then sngElapsedTime = 0.01
    lngPreviousFrame = (GetTickCount() And &H7FFFFFFF)
    
    Contador = Contador + sngElapsedTime
    
    
    If Efecto = 1 Then
    
    If Contador > 0.1 * 4.6 / timerElapsedTime And Iniciadas < ParticleCounts Then
        Iniciadas = Iniciadas + 1
        Iniciar (Iniciadas)
        Contador = 0
    End If
    
    For i = 0 To Iniciadas
        With Particles(i)
            Call .UpdateParticle(sngElapsedTime)
            
            '//If the particle is invisible, reset it again.
            If .sngA <= 0 Then

                Call Reset(i)
                
            End If
            If Matar Then
                Desa = Desa - sngElapsedTime / 1000
                If Desa < 0 Then Desa = 0
            End If
        
            vertsPoints(i).color = D3DColorMake(.sngR, .sngG, .sngB, .sngA * Desa)
            If i = Iniciadas Then
                vertsPoints(i).x = sngX + OffX
                vertsPoints(i).y = sngY + OffY
            Else
                vertsPoints(i).x = .sngX + OffX
                vertsPoints(i).y = .sngY + OffY
            End If
        End With
    Next i
    
    ElseIf Efecto = 2 Then
        With Particles(0)
            Call .UpdateParticle(sngElapsedTime)

            vertsPoints(0).color = D3DColorMake(.sngR, .sngG, .sngB, 1)
            vertsPoints(0).x = sngX + OffX
            vertsPoints(0).y = sngY + OffY

        End With
    End If
End Sub

Public Sub Render()
If Efecto = 1 Then
    With D3DDevice
        '//Set the render states for using point sprites
        .SetRenderState D3DRS_POINTSPRITE_ENABLE, 1 'True
        .SetRenderState D3DRS_POINTSCALE_ENABLE, 0 'True
        .SetRenderState D3DRS_POINTSIZE, lngFloatSize
        .SetRenderState D3DRS_POINTSIZE_MIN, lngFloat0
        .SetRenderState D3DRS_POINTSCALE_A, lngFloat0
        .SetRenderState D3DRS_POINTSCALE_B, lngFloat0
        .SetRenderState D3DRS_POINTSCALE_C, lngFloat1
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
        .SetRenderState D3DRS_ALPHABLENDENABLE, 1
        
        '//Set up the vertex shader
        .SetVertexShader D3DFVF_TLVERTEX
        
        '//Set our texture
        Call Engine_ReadyTexture(Texture)
        '.SetTexture 0, myTexture
        
        '//And draw all our particles :D
        .DrawPrimitiveUP D3DPT_POINTLIST, ParticleCounts, vertsPoints(0), Len(vertsPoints(0))
        
        '//Reset states back for normal rendering
        .SetRenderState D3DRS_ALPHABLENDENABLE, 1
        .SetRenderState D3DRS_POINTSPRITE_ENABLE, 1 'False
        .SetRenderState D3DRS_POINTSCALE_ENABLE, 0 'False
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End With
ElseIf Efecto = 2 Then
    Call Engine_Render_Rectangle(vertsPoints(0).x + 256, vertsPoints(0).y + 256, 32, 32, 0, 0, 32, 32, , , Angulo * 57, Texture)
End If
End Sub

Public Sub ReLocate(sngNewX As Single, sngNewY As Single)
    sngX = sngNewX
    sngY = sngNewY
End Sub
Public Sub Move(sngNewX As Single, sngNewY As Single)
    OffX = sngNewX
    OffY = sngNewY
End Sub

