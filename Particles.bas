Attribute VB_Name = "Particles"
Option Explicit

Private Type Effect
    X As Single                 'Location of effect
    Y As Single
    GoToX As Single             'Location to move to
    GoToY As Single
    KillWhenAtTarget As Boolean     'If the effect is at its target (GoToX/Y), then Progression is set to 0
    KillWhenTargetLost As Boolean   'Kill the effect if the target is lost (sets progression = 0)
    Gfx As Byte                 'Particle texture used
    Used As Boolean             'If the effect is in use
    EffectNum As Byte           'What number of effect that is used
    Modifier As Integer         'Misc variable (depends on the effect)
    FloatSize As Long           'The size of the particles
    Direction As Integer        'Misc variable (depends on the effect)
    Particles() As Particle     'Information on each particle
    Progression As Single       'Progression state, best to design where 0 = effect ends
    Looping As Boolean
    PartVertex() As TLVERTEX    'Used to point render particles
    PreviousFrame As Long       'Tick time of the last frame
    PreviousMove As Long
    NowMove As Long
    ParticleCount As Integer    'Number of particles total
    ParticlesLeft As Integer    'Number of particles left - only for non-repetitive effects
    BindToChar As Integer       'Setting this value will bind the effect to move towards the character
    BindSpeed As Single         'How fast the effect moves towards the character
    BoundToMap As Byte          'If the effect is bound to the map or not (used only by the map editor)
    TargetAA As Single
    R As Single
    G As Single
    B As Single
    A As Single
    Size As Byte
    SizeP As Single
    Ray As Byte
End Type

Public NumEffects As Integer   'Maximum number of effects at once
Public Effect() As Effect   'List of all the active effects



'Constants With The Order Number For Each Effect
Public Const EffectNum_Fire As Byte = 1             'Burn baby, burn! Flame from a central point that blows in a specified direction
Public Const EffectNum_Snow As Byte = 2             'Snow that covers the screen - weather effect
Public Const EffectNum_Heal As Byte = 3             'Healing effect that can bind to a character, ankhs float up and fade
Public Const EffectNum_Bless As Byte = 4            'Following three effects are same: create a circle around the central point
Public Const EffectNum_Protection As Byte = 5       ' (often the character) and makes the given particle on the perimeter
Public Const EffectNum_Strengthen As Byte = 6       ' which float up and fade out
Public Const EffectNum_Rain As Byte = 7             'Exact same as snow, but moves much faster and more alpha value - weather effect
Public Const EffectNum_EquationTemplate As Byte = 8 'Template for creating particle effects through equations - a page with some equations can be found here: http://www.vbgore.com/modules.php?name=Forums&file=viewtopic&t=221
Public Const EffectNum_Waterfall As Byte = 9        'Waterfall effect
Public Const EffectNum_Summon As Byte = 10
Public Const EffectNum_Explode As Byte = 11         'Explosion
Public Const EffectNum_Torch As Byte = 12           'Torch
Public Const EffectNum_Ray As Byte = 13             'Ray
Public Const EffectNum_Ice As Byte = 14             'Ice
Public Const EffectNum_Necro As Byte = 15           'Green Ray
Public Const EffectNum_Green As Byte = 16           'Green Explosion
Public Const EffectNum_Lissajous As Byte = 17       'Lissajous Curve (L)
Public Const EffectNum_Curse As Byte = 18       'Lissajous Curve (L)
Public Const EffectNum_Aura As Byte = 19
Public Const EffectNum_Aura2 As Byte = 20
Public Const EffectNum_Medit As Byte = 21
Public Const EffectNum_Atom As Byte = 22
Public Const EffectNum_Teleport As Byte = 23
Public Const EffectNum_Fountain As Byte = 24
Public Const EffectNum_Spell As Byte = 25
Public Const EffectNum_Light As Byte = 26
Public Const EffectNum_Smoke As Byte = 27
Public Const EffectNum_Misile As Byte = 28
Public Const EffectNum_Holy As Byte = 29
Public Const EffectNum_SmallTorch As Byte = 30
Public Const EffectNum_PortalGroso As Byte = 31
Public Const EffectNum_Nova As Byte = 32
Public Const EffectNum_Implode As Byte = 33
Public Const EffectNum_RedFountain As Byte = 34
Public Const EffectNum_MeditMAX As Byte = 21

Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef Destination As Any, ByVal Length As Long)

Function Effect_EquationTemplate_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Progression As Single = 1) As Integer
'*****************************************************************
'Particle effect template for effects as described on the
'wiki page: http://www.vbgore.com/Particle_effect_equations
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_EquationTemplate_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_EquationTemplate_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_EquationTemplate  'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True                     'Enable the effect
    Effect(EffectIndex).X = X                           'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx                       'Set the graphic
    Effect(EffectIndex).Progression = Progression       'If we loop the effect

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(18)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).Rhw = 1
        Effect_EquationTemplate_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_EquationTemplate_Reset(ByVal EffectIndex As Integer, ByVal index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_EquationTemplate_Reset
'*****************************************************************
Dim X As Single
Dim Y As Single
Dim R As Single
    
    Effect(EffectIndex).Progression = Effect(EffectIndex).Progression + 0.1
    'R = (index / 20) * exp(index / Effect(EffectIndex).Progression Mod 3)
    R = (index / 10) + (Effect(EffectIndex).Progression / ((Rnd * 0.3) + 0.7))
    X = R * Cos(index)
    Y = R * Sin(index)
    
    'Reset the particle
    Effect(EffectIndex).Particles(index).ResetIt Effect(EffectIndex).X + X, Effect(EffectIndex).Y + Y, 0, 0, 0, 0
    Effect(EffectIndex).Particles(index).ResetColor 1, 0.1 + (Rnd * 0.4), 0.2, 0.7, 0.4 + (Rnd * 0.2)

End Sub

Private Sub Effect_EquationTemplate_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_EquationTemplate_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    
    
    'Go Through The Particle Loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0.2 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression < 50 Then

                    'Reset the particle
                    Effect_EquationTemplate_Reset EffectIndex, LoopC

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).Color = 0

                End If

            Else

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Function Effect_Implode_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Progression As Single = 1) As Integer
'*****************************************************************
'Particle effect template for effects as described on the
'wiki page: http://www.vbgore.com/Particle_effect_equations
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_EquationTemplate_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Implode_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Implode  'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True                     'Enable the effect
    Effect(EffectIndex).X = X                           'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx                       'Set the graphic
    Effect(EffectIndex).Progression = Progression       'If we loop the effect

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(18)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).Rhw = 1
        Effect_Implode_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Implode_Reset(ByVal EffectIndex As Integer, ByVal index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_EquationTemplate_Reset
'*****************************************************************
Dim X As Single
Dim Y As Single
Dim R As Single
    
    Effect(EffectIndex).Progression = Effect(EffectIndex).Progression + 0.1
    'R = (index / 20) * exp(index / Effect(EffectIndex).Progression Mod 3)
    R = (91 - (index / 20) - (Effect(EffectIndex).Progression / ((Rnd * 0.1) + 0.9))) * 0.5
    X = R * Cos(index)
    Y = R * Sin(index) * 0.5
    
    'Reset the particle
    Effect(EffectIndex).Particles(index).ResetIt Effect(EffectIndex).X + X, Effect(EffectIndex).Y + Y, 0, 0, 0, 0
    Effect(EffectIndex).Particles(index).ResetColor 0.4 + (Rnd * 0.6), 0.4 + (Rnd * 0.6), 0.8, 0.7, 0.4 + (Rnd * 0.2)

End Sub

Private Sub Effect_Implode_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_EquationTemplate_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    
    
    'Go Through The Particle Loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0.2 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression < 50 Then

                    'Reset the particle
                    Effect_Implode_Reset EffectIndex, LoopC

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).Color = 0

                End If

            Else

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Function Effect_Nova_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Progression As Single = 1) As Integer
'*****************************************************************
'Particle effect template for effects as described on the
'wiki page: http://www.vbgore.com/Particle_effect_equations
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_EquationTemplate_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Nova_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Nova  'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True                     'Enable the effect
    Effect(EffectIndex).X = X                           'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx                       'Set the graphic
    Effect(EffectIndex).Progression = Progression       'If we loop the effect

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(24)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).Rhw = 1
        Effect_Nova_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Nova_Reset(ByVal EffectIndex As Integer, ByVal index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_EquationTemplate_Reset
'*****************************************************************
Dim X As Single
Dim Y As Single
Dim R As Single
    
    Effect(EffectIndex).Progression = Effect(EffectIndex).Progression + 0.05
    'R = (index / 20) * exp(index / Effect(EffectIndex).Progression Mod 3)
    R = (index / 10) + (Effect(EffectIndex).Progression / ((Rnd * 0.3) + 0.7))
    X = R * Cos(index) * 2
    Y = R * Sin(index) * 0.5
    
    'Reset the particle
    Effect(EffectIndex).Particles(index).ResetIt Effect(EffectIndex).X + X, Effect(EffectIndex).Y + Y, 0, 0, 0, 0
    Effect(EffectIndex).Particles(index).ResetColor 1, 0.25 + (Rnd * 0.6), 0.2, 0.7, 0.2 + (Rnd * 0.2)

End Sub

Private Sub Effect_Nova_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_EquationTemplate_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    
    
    'Go Through The Particle Loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0.3 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression < 50 Then

                    'Reset the particle
                    Effect_Nova_Reset EffectIndex, LoopC

                Else
                    
                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).Color = 0

                End If

            Else

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Function Effect_Aura_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Progression As Single = 1) As Integer
Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Aura_Begin = EffectIndex
    
    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Aura  'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True                     'Enable the effect
    Effect(EffectIndex).X = X                           'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx                       'Set the graphic
    Effect(EffectIndex).Progression = Progression       'If we loop the effect

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(3)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).Rhw = 1
        Effect_Aura_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Aura_Reset(ByVal EffectIndex As Integer, ByVal index As Long)
Dim X As Single
Dim Y As Single
Dim R As Single
Dim ind As Integer
    Effect(EffectIndex).Progression = Effect(EffectIndex).Progression + 0.1
    ind = CInt(index / 10) * 10
    R = ((index + 100) / 4) * exp((index + 100) / 2000)
    X = R * Cos(index) * 0.3 * 0.25
    Y = R * Sin(index) * 0.2 * 0.25
    'Reset the particle
    If Rnd * 20 < 1 Then
        Effect(EffectIndex).Particles(index).ResetIt Effect(EffectIndex).X + X, Effect(EffectIndex).Y + Y, 0, 0, 0, -1.5 * (ind / Effect(EffectIndex).ParticleCount)
    Else
        Effect(EffectIndex).Particles(index).ResetIt Effect(EffectIndex).X + X, Effect(EffectIndex).Y + Y, 0, 0, 0, 0
    End If
    Effect(EffectIndex).Particles(index).ResetColor 1, (1 * ind / Effect(EffectIndex).ParticleCount), 0, 1, IIf(ind / Effect(EffectIndex).ParticleCount / 7 < 0.03, 0.03, ind / Effect(EffectIndex).ParticleCount / 7)

End Sub

Private Sub Effect_Aura_Update(ByVal EffectIndex As Integer)
Dim ElapsedTime As Single
Dim LoopC As Long
Dim Owner As Integer
    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime
    
    For LoopC = 1 To LastChar
        If EffectIndex = CharList(LoopC).AuraIndex Then
            Owner = LoopC
        End If
    Next
    
    If ClientSetup.bGraphics < 2 Then Effect(EffectIndex).Used = False
    
    If Owner = 0 Then
        Effect(EffectIndex).Used = False
    End If
    
    'Go Through The Particle Loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update The Particle
            If EffectIndex <> CharList(UserCharIndex).AuraIndex Then
                Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime
            Else
                Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime, True
            End If
            
            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Then

                    'Reset the particle
                    Effect_Aura_Reset EffectIndex, LoopC

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).Color = 0

                End If

            Else

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub
Function Effect_Aura2_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Progression As Single = 1) As Integer
Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Aura2_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Aura2  'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True                     'Enable the effect
    Effect(EffectIndex).X = X                           'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx                       'Set the graphic
    Effect(EffectIndex).Progression = Progression       'If we loop the effect

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(3)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).Rhw = 1
        Effect_Aura2_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Aura2_Reset(ByVal EffectIndex As Integer, ByVal index As Long)
Dim X As Single
Dim Y As Single
Dim R As Single
Dim ind As Integer
    Effect(EffectIndex).Progression = Effect(EffectIndex).Progression + 0.1
    ind = CInt(index / 10) * 10
    R = ((index + 100) / 4) * Tan(ind / 950)  'Exp((Index + 100) / 2000)
    X = R * Cos(index) * 0.3 * 0.25
    Y = R * Sin(index) * 0.2 * 0.25
    'Reset the particle
    If Rnd * 20 < 1 Then
        Effect(EffectIndex).Particles(index).ResetIt Effect(EffectIndex).X + X, Effect(EffectIndex).Y + Y, 0, 0, 0, -1.5 * (ind / Effect(EffectIndex).ParticleCount)
    Else
        Effect(EffectIndex).Particles(index).ResetIt Effect(EffectIndex).X + X, Effect(EffectIndex).Y + Y, 0, 0, 0, 0
    End If
    Effect(EffectIndex).Particles(index).ResetColor (1 * ind / Effect(EffectIndex).ParticleCount), (1 * ind / Effect(EffectIndex).ParticleCount), 1, 1, IIf(ind / Effect(EffectIndex).ParticleCount / 7 < 0.03, 0.03, ind / Effect(EffectIndex).ParticleCount / 7)

End Sub

Private Sub Effect_Aura2_Update(ByVal EffectIndex As Integer)
Dim ElapsedTime As Single
Dim LoopC As Long
Dim Owner As Integer
    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime
    
    For LoopC = 1 To LastChar
        If EffectIndex = CharList(LoopC).AuraIndex Then
            Owner = LoopC
        End If
    Next
    
    If ClientSetup.bGraphics < 2 Then Effect(EffectIndex).Used = False
    
    If Owner = 0 Then
        Effect(EffectIndex).Used = False
    End If
    
    'Go Through The Particle Loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update The Particle
            If EffectIndex <> CharList(UserCharIndex).AuraIndex Then
                Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime
            Else
                Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime, True
            End If
            
            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Then

                    'Reset the particle
                    Effect_Aura2_Reset EffectIndex, LoopC

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).Color = 0

                End If

            Else

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Function Effect_PortalGroso_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Progression As Single = 1) As Integer
Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_PortalGroso_Begin = EffectIndex
    
    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_PortalGroso  'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True                     'Enable the effect
    Effect(EffectIndex).X = X                           'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx                       'Set the graphic
    Effect(EffectIndex).Progression = Progression       'If we loop the effect

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(20)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).Rhw = 1
        Effect_PortalGroso_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_PortalGroso_Reset(ByVal EffectIndex As Integer, ByVal index As Long)
Dim X As Single
Dim Y As Single
Dim R As Single
Dim ind As Integer
    Effect(EffectIndex).Progression = Effect(EffectIndex).Progression + 0.1
    ind = CInt(index / 10) * 10
    R = ((index + 100) / 4) * exp((index + 100) / 2000)
    X = R * Cos(index) * 0.25 '* 0.3 * 0.25
    Y = R * Sin(index) * 0.25 '* 0.2 * 0.25
    'Reset the particle
    'If Rnd * 20 < 1 Then
    '    Effect(EffectIndex).Particles(index).ResetIt Effect(EffectIndex).x + x, Effect(EffectIndex).y + y, 0, 0, 0, -1.5 * (ind / Effect(EffectIndex).ParticleCount)
    'Else
        Effect(EffectIndex).Particles(index).ResetIt Effect(EffectIndex).X + X, Effect(EffectIndex).Y + Y, 0, 0, 0, 0
    'End If
    Effect(EffectIndex).Particles(index).ResetColor 0.2, 0, (0.7 * ind / Effect(EffectIndex).ParticleCount), 1, IIf(ind / Effect(EffectIndex).ParticleCount / 7 < 0.03, 0.03, ind / Effect(EffectIndex).ParticleCount / 7)

End Sub

Private Sub Effect_PortalGroso_Update(ByVal EffectIndex As Integer)
Dim ElapsedTime As Single
Dim LoopC As Long
Dim Owner As Integer
    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime
    
    'For LoopC = 1 To LastChar
    '    If EffectIndex = CharList(LoopC).AuraIndex Then
    '        Owner = LoopC
    '    End If
    'Next
    
    'If ClientSetup.bGraphics < 2 Then Effect(EffectIndex).Used = False
    
    'If Owner = 0 Then
    '    Effect(EffectIndex).Used = False
    'End If
    
    'Go Through The Particle Loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            ''Update The Particle
            'If EffectIndex <> CharList(UserCharIndex).AuraIndex Then
                Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime
            'Else
             '   Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime, True
            'End If
            
            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Then

                    'Reset the particle
                    Effect_PortalGroso_Reset EffectIndex, LoopC

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).Color = 0

                End If

            Else

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Function Effect_Teleport_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Size As Byte = 30, Optional ByVal Time As Single = 10) As Integer
Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Teleport_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Teleport     'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles - Particles * 0.25 * (2 - ClientSetup.bGraphics)       'Set the number of particles
    Effect(EffectIndex).Used = True             'Enabled the effect
    Effect(EffectIndex).X = X                   'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                   'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx               'Set the graphic
    Effect(EffectIndex).Modifier = Size         'How large the circle is
    Effect(EffectIndex).Progression = Time      'How long the effect will last

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(15)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).Rhw = 1
        Effect_Teleport_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Teleport_Reset(ByVal EffectIndex As Integer, ByVal index As Long)
Dim A As Single
Dim X As Single
Dim Y As Single


    If Rnd * 10 < 5 Then
        'Get the positions
        A = Rnd * 360 * DegreeToRadian
        X = Effect(EffectIndex).X - (Sin(A) * Effect(EffectIndex).Modifier) / 2.2 '* (0.8 + Rnd * 0.2)
        Y = Effect(EffectIndex).Y + (Cos(A) * Effect(EffectIndex).Modifier)
    Else
        A = Rnd * 360 * DegreeToRadian
        X = Effect(EffectIndex).X - (Sin(A) * Effect(EffectIndex).Modifier / 2) / 2.2 '* (0.8 + Rnd * 0.2)
        Y = Effect(EffectIndex).Y + (Cos(A) * Effect(EffectIndex).Modifier / 2)
    End If
    'Reset the particle
    Effect(EffectIndex).Particles(index).ResetIt X, Y, 0, Rnd * -1, 0, -2
    Effect(EffectIndex).Particles(index).ResetColor 1, Rnd * 0.5, Rnd * 0.5, 0.6 + (Rnd * 0.4), 0.2 + (Rnd * 0.2)

End Sub

Private Sub Effect_Teleport_Update(ByVal EffectIndex As Integer)
Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Update the life span
    'If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime

    'Go Through The Particle Loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Then

                    'Reset the particle
                    Effect_Teleport_Reset EffectIndex, LoopC

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).Color = 0

                End If

            Else

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub
Function Effect_Medit_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Size As Byte = 30, Optional ByVal Time As Single = 10, Optional R As Single) As Integer
Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Medit_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Medit    'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles - Particles * 0.25 * (2 - ClientSetup.bGraphics)   'Set the number of particles
    Effect(EffectIndex).Used = True             'Enabled the effect
    Effect(EffectIndex).X = X                   'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                   'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx               'Set the graphic
    Effect(EffectIndex).Modifier = Size         'How large the circle is
    Effect(EffectIndex).Progression = Time      'How long the effect will last
    
    Effect(EffectIndex).R = R

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(20)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).Rhw = 1
        Effect_Medit_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Medit_Reset(ByVal EffectIndex As Integer, ByVal index As Long)
Dim A As Single
Dim X As Single
Dim Y As Single

    'Get the positions
    A = Rnd * 360 * DegreeToRadian
    
    Do While A > 3 And A < 3.27
        Randomize 1000
        A = Rnd * 60 * DegreeToRadian * RandomNumber(1, 6)
    Loop
    
    X = Effect(EffectIndex).X - (Sin(A) * Effect(EffectIndex).Modifier)
    Y = Effect(EffectIndex).Y + (Cos(A) * Effect(EffectIndex).Modifier) / 2

    'Reset the particle
    Select Case Effect(EffectIndex).R
        Case 0
            Effect(EffectIndex).Particles(index).ResetIt X, Y, 0, Rnd * -1, 0, -0.5
            Effect(EffectIndex).Particles(index).ResetColor 0.9, 0.1, 0.1, 0.7 + (Rnd * 0.3), 0.05 + (Rnd * 0.1)
        Case 1
            Effect(EffectIndex).Particles(index).ResetIt X, Y, 0, Rnd * -1, 0, -0.9
            Effect(EffectIndex).Particles(index).ResetColor 0.1, 0.9, 0.1, 0.7 + (Rnd * 0.3), 0.05 + (Rnd * 0.1)
        Case 2
            Effect(EffectIndex).Particles(index).ResetIt X, Y, 0, Rnd * -1, 0, -1.4
            Effect(EffectIndex).Particles(index).ResetColor 0.1, 0.9, 0.9, 0.7 + (Rnd * 0.3), 0.05 + (Rnd * 0.1)
        Case 3
            Effect(EffectIndex).Particles(index).ResetIt X, Y, 0, Rnd * -1, 0, -1.6
            Effect(EffectIndex).Particles(index).ResetColor 0.9, 0.9, 0.6, 0.7 + (Rnd * 0.3), 0.05 + (Rnd * 0.1)
        Case 4
            Effect(EffectIndex).Particles(index).ResetIt X, Y, 0, Rnd * -1, 0, -1.8
            Effect(EffectIndex).Particles(index).ResetColor 0.9 * Rnd, 0.9 * Rnd, 0.7, 0.7 + (Rnd * 0.3), 0.05 + (Rnd * 0.1)
        Case 5
        Case 6
    End Select
End Sub
Public Sub Effect_Medit_Update(ByVal EffectIndex As Integer)
Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Update the life span
    'If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime

    'Go through the particle loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Then

                    'Reset the particle
                    Effect_Medit_Reset EffectIndex, LoopC

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).Color = 0

                End If

            Else

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Function Effect_MeditMAX_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Size As Byte = 30, Optional ByVal Time As Single = 10, Optional R As Single) As Integer
Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_MeditMAX_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_MeditMAX    'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles - Particles * 0.25 * (2 - ClientSetup.bGraphics)   'Set the number of particles
    Effect(EffectIndex).Used = True             'Enabled the effect
    Effect(EffectIndex).X = X                   'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                   'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx               'Set the graphic
    Effect(EffectIndex).Modifier = Size         'How large the circle is
    Effect(EffectIndex).Progression = Time      'How long the effect will last
    
    Effect(EffectIndex).R = R

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(20)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).Rhw = 1
        Effect_MeditMAX_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_MeditMAX_Reset(ByVal EffectIndex As Integer, ByVal index As Long)
Dim A As Single
Dim X As Single
Dim Y As Single
Dim AccX As Single

    'Get the positions
    A = Rnd * 360 * DegreeToRadian
    
    Do While A > 3 And A < 3.27
        Randomize 1000
        A = Rnd * 60 * DegreeToRadian * RandomNumber(1, 6)
    Loop
    
    X = Effect(EffectIndex).X - (Sin(A) * Effect(EffectIndex).Modifier)
    Y = Effect(EffectIndex).Y + (Cos(A) * Effect(EffectIndex).Modifier) / 2

    
    Effect(EffectIndex).Particles(index).ResetIt X, Y, -(Sgn(Sin(A)) * 4 * (Rnd - 0.2)), Rnd * -1, 0, -1.8
    Effect(EffectIndex).Particles(index).ResetColor 0.9, 0.9, 0.7 * Rnd, 0.7 + (Rnd * 0.3), 0.05 + (Rnd * 0.1)

End Sub
Public Sub Effect_MeditMAX_Update(ByVal EffectIndex As Integer)
Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Update the life span
    'If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime

    'Go through the particle loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Then

                    'Reset the particle
                    Effect_MeditMAX_Reset EffectIndex, LoopC

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).Color = 0

                End If

            Else

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub



Function Effect_Atom_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Size As Byte = 30, Optional ByVal Time As Single = 10) As Integer

Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Atom_Begin = EffectIndex

    'Set the effect's variables
    Effect(EffectIndex).EffectNum = EffectNum_Atom    'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles           'Set the number of particles
    Effect(EffectIndex).Used = True             'Enabled the effect
    Effect(EffectIndex).X = X                   'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                   'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx               'Set the graphic
    Effect(EffectIndex).Modifier = Size         'How large the circle is
    Effect(EffectIndex).Progression = Time      'How long the effect will last

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(16)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).Rhw = 1
        Effect_Atom_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Atom_Reset(ByVal EffectIndex As Integer, ByVal index As Long)

Dim A As Single
Dim X As Single
Dim Y As Single
Dim R As Single
    'Get the positions
    A = Rnd * 360 * DegreeToRadian
    R = Rnd * 4
    If R < 1 Then
        X = Effect(EffectIndex).X - (Sin(A) * Effect(EffectIndex).Modifier) / 3 + (Cos(A) * Effect(EffectIndex).Modifier)
        Y = Effect(EffectIndex).Y + (Cos(A) * Effect(EffectIndex).Modifier)
        Effect(EffectIndex).Particles(index).ResetColor 0.2, 1, 0.2, 0.6 + (Rnd * 0.4), 0.06 + (Rnd * 0.2)
    ElseIf R < 2 Then
        X = Effect(EffectIndex).X - (Sin(A) * Effect(EffectIndex).Modifier)
        Y = Effect(EffectIndex).Y + (Cos(A) * Effect(EffectIndex).Modifier) / 3 + (Sin(A) * Effect(EffectIndex).Modifier)
        Effect(EffectIndex).Particles(index).ResetColor 1, 1, 0.2, 0.6 + (Rnd * 0.4), 0.06 + (Rnd * 0.2)
    ElseIf R < 3 Then
        X = Effect(EffectIndex).X - (Sin(A) * Effect(EffectIndex).Modifier) / 3
        Y = Effect(EffectIndex).Y + (Cos(A) * Effect(EffectIndex).Modifier)
        Effect(EffectIndex).Particles(index).ResetColor 1, 0.2, 0.2, 0.6 + (Rnd * 0.4), 0.06 + (Rnd * 0.2)
    ElseIf R < 4 Then
        X = Effect(EffectIndex).X - (Sin(A) * Effect(EffectIndex).Modifier)
        Y = Effect(EffectIndex).Y + (Cos(A) * Effect(EffectIndex).Modifier) / 3
        
        Effect(EffectIndex).Particles(index).ResetColor 0.2, 0.2, 1, 0.6 + (Rnd * 0.4), 0.06 + (Rnd * 0.2)
    End If
    
    'Reset the particle
    Effect(EffectIndex).Particles(index).ResetIt X, Y, 0, 0, 0, -1
    

End Sub

Private Sub Effect_Atom_Update(ByVal EffectIndex As Integer)
Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate the time difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime

    'Go through the particle loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check if particle is in use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update the particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Then

                    'Reset the particle
                    Effect_Atom_Reset EffectIndex, LoopC

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).Color = 0

                End If

            Else

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub


Function Effect_Fountain_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer) As Integer

Dim EffectIndex As Integer
Dim LoopC As Long

'Get the next open effect slot

    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Fountain_Begin = EffectIndex

    'Set the effect's variables
    Effect(EffectIndex).EffectNum = EffectNum_Fountain     'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles - Particles * 0.25 * (2 - ClientSetup.bGraphics)           'Set the number of particles
    Effect(EffectIndex).Used = True             'Enabled the effect
    Effect(EffectIndex).X = X                   'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                   'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx               'Set the graphic

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(20)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).Rhw = 1
        Effect_Fountain_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Fountain_Reset(ByVal EffectIndex As Integer, ByVal index As Long)

    If Int(Rnd * 10) < 5 Then
        Effect(EffectIndex).Particles(index).ResetIt Effect(EffectIndex).X + (Rnd * 10) - 5, Effect(EffectIndex).Y - (Rnd * 10), -1 + (Rnd * -5), -15 - (Rnd * 3), 0, 1.1 + Rnd * 0.1
    Else
        Effect(EffectIndex).Particles(index).ResetIt Effect(EffectIndex).X + (Rnd * 10) - 5, Effect(EffectIndex).Y - (Rnd * 10), 1 + (Rnd * 5), -15 - (Rnd * 3), 0, 1.1 + Rnd * 0.1
    End If
    Effect(EffectIndex).Particles(index).ResetColor 0.1, 0.1, 0.9, 0.6 + (Rnd * 0.4), 0.02 + Rnd * 0.05
    
End Sub

Private Sub Effect_Fountain_Update(ByVal EffectIndex As Integer)
Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime

    'Go through the particle loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
    
        With Effect(EffectIndex).Particles(LoopC)
    
            'Check if the particle is in use
            If .Used Then
    
                'Update The Particle
                .UpdateParticle ElapsedTime

                'Check if the particle is ready to die
                If (.sngA < 0) Or (.sngY > Effect(EffectIndex).Y + 100) Then
    
                    'Reset the particle
                    Effect_Fountain_Reset EffectIndex, LoopC
    
                Else

                    'Set the particle information on the particle vertex
                    Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(.sngR, .sngG, .sngB, .sngA)
                    Effect(EffectIndex).PartVertex(LoopC).X = .sngX
                    Effect(EffectIndex).PartVertex(LoopC).Y = .sngY
    
                End If
    
            End If
            
        End With

    Next LoopC

End Sub

Function Effect_RedFountain_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer) As Integer

Dim EffectIndex As Integer
Dim LoopC As Long

'Get the next open effect slot

    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_RedFountain_Begin = EffectIndex

    'Set the effect's variables
    Effect(EffectIndex).EffectNum = EffectNum_RedFountain     'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles - Particles * 0.25 * (2 - ClientSetup.bGraphics)           'Set the number of particles
    Effect(EffectIndex).Used = True             'Enabled the effect
    Effect(EffectIndex).X = X                   'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                   'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx               'Set the graphic

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(20)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).Rhw = 1
        Effect_RedFountain_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_RedFountain_Reset(ByVal EffectIndex As Integer, ByVal index As Long)

    'If Int(Rnd * 10) < 6 Then
        Effect(EffectIndex).Particles(index).ResetIt Effect(EffectIndex).X + (Rnd * 10) - 5, Effect(EffectIndex).Y - (Rnd * 10), 0, 1, 0, -1 - Rnd * 0.25
    'Else
        'Effect(EffectIndex).Particles(index).ResetIt Effect(EffectIndex).x + (Rnd * 10) - 5, Effect(EffectIndex).y - (Rnd * 10), 1 + (Rnd * 5), -15 - (Rnd * 3), 0, 1.1 + Rnd * 0.1
    'End If
    Effect(EffectIndex).Particles(index).ResetColor 0.9, Rnd * 0.7, 0.1, 0.6 + (Rnd * 0.4), 0.035 + Rnd * 0.01
    
End Sub

Private Sub Effect_RedFountain_Update(ByVal EffectIndex As Integer)
Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime

    'Go through the particle loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
    
        With Effect(EffectIndex).Particles(LoopC)
    
            'Check if the particle is in use
            If .Used Then
    
                'Update The Particle
                .UpdateParticle ElapsedTime

                'Check if the particle is ready to die
                If (.sngA < 0) Or (.sngY > Effect(EffectIndex).Y + 100) Then
    
                    'Reset the particle
                    Effect_RedFountain_Reset EffectIndex, LoopC
    
                Else

                    'Set the particle information on the particle vertex
                    Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(.sngR, .sngG, .sngB, .sngA)
                    Effect(EffectIndex).PartVertex(LoopC).X = .sngX
                    Effect(EffectIndex).PartVertex(LoopC).Y = .sngY
    
                End If
    
            End If
            
        End With

    Next LoopC

End Sub

Function Effect_Bless_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Size As Byte = 30, Optional ByVal Time As Single = 10) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Bless_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Bless_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Bless     'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles - Particles * 0.25 * (2 - ClientSetup.bGraphics)       'Set the number of particles
    Effect(EffectIndex).Used = True             'Enabled the effect
    Effect(EffectIndex).X = X                   'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                   'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx               'Set the graphic
    Effect(EffectIndex).Modifier = Size         'How large the circle is
    Effect(EffectIndex).Progression = Time      'How long the effect will last

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(20)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).Rhw = 1
        Effect_Bless_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Bless_Reset(ByVal EffectIndex As Integer, ByVal index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Bless_Reset
'*****************************************************************
Dim A As Single
Dim X As Single
Dim Y As Single

    'Get the positions
    A = Rnd * 360 * DegreeToRadian
    X = Effect(EffectIndex).X - (Sin(A) * Effect(EffectIndex).Modifier)
    Y = Effect(EffectIndex).Y + (Cos(A) * Effect(EffectIndex).Modifier)

    'Reset the particle
    Effect(EffectIndex).Particles(index).ResetIt X, Y, 0, Rnd * -1, 0, -2
    Effect(EffectIndex).Particles(index).ResetColor 1, 1, 0.2, 0.6 + (Rnd * 0.4), 0.06 + (Rnd * 0.2)

End Sub

Private Sub Effect_Bless_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Bless_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime

    
    'Go Through The Particle Loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Then

                    'Reset the particle
                    Effect_Bless_Reset EffectIndex, LoopC

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).Color = 0

                End If

            Else

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub


Function Effect_Explode_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Size As Byte = 30, Optional ByVal Time As Single = 10) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Bless_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Explode_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Explode     'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True             'Enabled the effect
    Effect(EffectIndex).X = X                   'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                   'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx               'Set the graphic
    Effect(EffectIndex).Modifier = Size         'How large the circle is
    Effect(EffectIndex).Progression = Time      'How long the effect will last

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(20)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).Rhw = 1
        'Effect_Explode_Reset EffectIndex, LoopC
    Next LoopC
    
    
    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime
    Effect_Explode_Update EffectIndex

End Function

Private Sub Effect_Explode_Reset(ByVal EffectIndex As Integer, ByVal index As Long)

Dim A As Single
Dim X As Single
Dim Y As Single


    'Get the positions
    A = Round(Rnd, 1) * 360 * DegreeToRadian
    
    
    X = Effect(EffectIndex).X - (Sin(A) * Effect(EffectIndex).Modifier) '+ index / 20
    Y = Effect(EffectIndex).Y + (Cos(A) * Effect(EffectIndex).Modifier) '+ index / 20
    

    'Reset the particle
    Effect(EffectIndex).Particles(index).ResetIt X, Y, 0, 0, 0, 0
    Effect(EffectIndex).Particles(index).ResetColor 1, 0.2 + (Rnd * 0.3), 0.2, 0.6 + (Rnd * 0.4), 0.06 + (Rnd * 0.2)

End Sub

Private Sub Effect_Explode_Update(ByVal EffectIndex As Integer)

Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime

    Effect(EffectIndex).Modifier = Effect(EffectIndex).Modifier + 4
    'Go Through The Particle Loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Then

                    'Reset the particle
                    Effect_Explode_Reset EffectIndex, LoopC

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).Color = 0

                End If

            Else

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Function Effect_Fire_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Direction As Integer = 180, Optional ByVal Progression As Single = 1) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Fire_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Fire      'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles - Particles * 0.25 * (2 - ClientSetup.bGraphics)       'Set the number of particles
    Effect(EffectIndex).Used = True     'Enabled the effect
    Effect(EffectIndex).X = X          'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx       'Set the graphic
    Effect(EffectIndex).Direction = Direction       'The direction the effect is animat
    Effect(EffectIndex).Progression = Progression   'Loop the effect

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(15)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).Rhw = 1
        Effect_Fire_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Fire_Reset(ByVal EffectIndex As Integer, ByVal index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Reset
'*****************************************************************

    'Reset the particle
    Effect(EffectIndex).Particles(index).ResetIt Effect(EffectIndex).X - 10 + Rnd * 20, Effect(EffectIndex).Y - 10 + Rnd * 20, -Sin((Effect(EffectIndex).Direction + (Rnd * 70) - 35) * DegreeToRadian) * 8, Cos((Effect(EffectIndex).Direction + (Rnd * 70) - 35) * DegreeToRadian) * 8, 0, 0
    Effect(EffectIndex).Particles(index).ResetColor 1, 0.1 + (Rnd * 0.4), 0.2, 0.4 + (Rnd * 0.2), 0.03 + (Rnd * 0.07)

End Sub

Private Sub Effect_Fire_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime
    
    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime
    
    'Go Through The Particle Loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Or Effect(EffectIndex).Progression = -5000 Then

                    'Reset the particle
                    Effect_Fire_Reset EffectIndex, LoopC
                    
                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).Color = 0

                End If

            Else

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Function Effect_Misile_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Direction As Integer = 180, Optional ByVal Progression As Single = 1) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Misile_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Misile      'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles - Particles * 0.25 * (2 - ClientSetup.bGraphics)       'Set the number of particles
    Effect(EffectIndex).Used = True     'Enabled the effect
    Effect(EffectIndex).X = X          'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx       'Set the graphic
    Effect(EffectIndex).Direction = Direction       'The direction the effect is animat
    Effect(EffectIndex).Progression = -5000   'Loop the effect

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(15)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).Rhw = 1
        'Effect_Misile_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Misile_Reset(ByVal EffectIndex As Integer, ByVal index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Reset
'*****************************************************************
    Dim Angle As Single
    Dim Color As Single
    Color = (Rnd * 0.4)
    If Effect(EffectIndex).TargetAA = 0 Then Effect(EffectIndex).TargetAA = Engine_GetAngle(Effect(EffectIndex).X, Effect(EffectIndex).Y, Effect(EffectIndex).GoToX, Effect(EffectIndex).GoToY)
    'Reset the particle
    Effect(EffectIndex).Particles(index).ResetIt Effect(EffectIndex).X - 10 + Rnd * 20, Effect(EffectIndex).Y - 10 + Rnd * 20, -Sin((Effect(EffectIndex).TargetAA + (Rnd * 50) - 35) * DegreeToRadian) * 8, Cos((Effect(EffectIndex).TargetAA + (Rnd * 50) - 35) * DegreeToRadian) * 8, 0, 0
    Effect(EffectIndex).Particles(index).ResetColor 0.5 + Color, 0.5 + Color, 0.5 + Color, 0.4 + (Rnd * 0.2), 0.2 + (Rnd * 0.07)

End Sub

Private Sub Effect_Misile_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime
    
    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime
    
    'Go Through The Particle Loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0 Or Not Engine_RectDistance(Effect(EffectIndex).X, Effect(EffectIndex).Y, Effect(EffectIndex).Particles(LoopC).sngX, Effect(EffectIndex).Particles(LoopC).sngY, 32, 32) Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Or Effect(EffectIndex).Progression = -5000 Then

                    'Reset the particle
                    Effect_Misile_Reset EffectIndex, LoopC
                    
                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).Color = 0

                End If

            Else

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Function Effect_Torch_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Direction As Integer = 180, Optional ByVal Progression As Single = 1) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Torch_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Torch     'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles - Particles * 0.25 * (2 - ClientSetup.bGraphics)       'Set the number of particles
    Effect(EffectIndex).Used = True     'Enabled the effect
    Effect(EffectIndex).X = X          'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx       'Set the graphic
    Effect(EffectIndex).Direction = Direction       'The direction the effect is animat
    Effect(EffectIndex).Progression = Progression   'Loop the effect

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(15)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).Rhw = 1
        Effect_Torch_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Torch_Reset(ByVal EffectIndex As Integer, ByVal index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Reset
'*****************************************************************

    'Reset the particle
    Effect(EffectIndex).Particles(index).ResetIt Effect(EffectIndex).X - 10 + Rnd * 20, Effect(EffectIndex).Y - 10 + Rnd * 20, Rnd * 2 * RandomNumber(-1, 1), Rnd * 2 * RandomNumber(-1, 1), 0, 0
    Effect(EffectIndex).Particles(index).ResetColor 1, 0.1 + (Rnd * 0.4), 0.2, 0.4 + (Rnd * 0.2), 0.1 + (Rnd * 0.07)

End Sub

Private Sub Effect_Torch_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime
    
    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime
    
    'Go Through The Particle Loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0.3 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Or Effect(EffectIndex).Progression = -5000 Then
                    
                    
                    'Reset the particle
                    Effect_Torch_Reset EffectIndex, LoopC
                    
                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).Color = 0

                End If

            Else
                
                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Function Effect_Light_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Direction As Integer = 180, Optional ByVal Progression As Single = 1) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Light_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Light     'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles - Particles * 0.25 * (2 - ClientSetup.bGraphics)      'Set the number of particles
    Effect(EffectIndex).Used = True     'Enabled the effect
    Effect(EffectIndex).X = X          'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx       'Set the graphic
    Effect(EffectIndex).Direction = Direction       'The direction the effect is animat
    Effect(EffectIndex).Progression = Progression   'Loop the effect

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(15)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).Rhw = 1
        Effect_Light_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Light_Reset(ByVal EffectIndex As Integer, ByVal index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Reset
'*****************************************************************
    Dim v As Single
    v = Rnd * 15
    'Reset the particle
    Effect(EffectIndex).Particles(index).ResetIt Effect(EffectIndex).X - 7.5 + v, Effect(EffectIndex).Y - 7.5 + Rnd * 15, 0, -1, 0, Rnd * -3
    Effect(EffectIndex).Particles(index).ResetColor 0.9, Rnd * 0.7, 0.1, 0.6 + (Rnd * 0.2), 0.03 + (Rnd * 0.07)

End Sub

Private Sub Effect_Light_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime
    
    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime
    
    'Go Through The Particle Loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0.3 Or Effect(EffectIndex).Particles(LoopC).sngY + 40 < Effect(EffectIndex).Y Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Or Effect(EffectIndex).Progression = -5000 Then
                    
                    
                    'Reset the particle
                    Effect_Light_Reset EffectIndex, LoopC
                    
                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).Color = 0

                End If

            Else
                
                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Function Effect_Holy_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Direction As Integer = 180, Optional ByVal Progression As Single = 1) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Holy_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Holy     'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles - Particles * 0.25 * (2 - ClientSetup.bGraphics)      'Set the number of particles
    Effect(EffectIndex).Used = True     'Enabled the effect
    Effect(EffectIndex).X = X          'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx       'Set the graphic
    Effect(EffectIndex).Direction = Direction       'The direction the effect is animat
    Effect(EffectIndex).Progression = Progression   'Loop the effect

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(24)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).Rhw = 1
        Effect_Holy_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Holy_Reset(ByVal EffectIndex As Integer, ByVal index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Reset
'*****************************************************************
    Dim v As Integer
    v = RandomNumber(1, 4)
    'Reset the particle
    Select Case v
        Case 1
            Effect(EffectIndex).Particles(index).ResetIt Effect(EffectIndex).X, Effect(EffectIndex).Y, 2 + Rnd * 1, 0, 0, 0, Effect(EffectIndex).X, Effect(EffectIndex).Y
        Case 2
            Effect(EffectIndex).Particles(index).ResetIt Effect(EffectIndex).X, Effect(EffectIndex).Y, 0, 2 + Rnd * 1, 0, 0, Effect(EffectIndex).X, Effect(EffectIndex).Y
        Case 3
            Effect(EffectIndex).Particles(index).ResetIt Effect(EffectIndex).X, Effect(EffectIndex).Y, -2 - Rnd * 1, 0, 0, 0, Effect(EffectIndex).X, Effect(EffectIndex).Y
        Case 4
            Effect(EffectIndex).Particles(index).ResetIt Effect(EffectIndex).X, Effect(EffectIndex).Y, 0, -2 - Rnd * 1, 0, 0, Effect(EffectIndex).X, Effect(EffectIndex).Y
    End Select
    
    Effect(EffectIndex).Particles(index).ResetColor 0.8 + Rnd * 0.1, 0.8 + Rnd * 0.1, 0.6 + (Rnd * 0.2), 0.6 + (Rnd * 0.2), 0.02 + (Rnd * 0.07)

End Sub

Private Sub Effect_Holy_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime
    
    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime
    
    'Go Through The Particle Loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0.3 Then 'Or Effect(EffectIndex).Particles(LoopC).sngY + 40 < Effect(EffectIndex).y Or Effect(EffectIndex).Particles(LoopC).sngY - 40 > Effect(EffectIndex).y Or Effect(EffectIndex).Particles(LoopC).sngX - 40 > Effect(EffectIndex).x Or Effect(EffectIndex).Particles(LoopC).sngX + 40 < Effect(EffectIndex).x Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Or Effect(EffectIndex).Progression = -5000 Then
                    
                    
                    'Reset the particle
                    Effect_Holy_Reset EffectIndex, LoopC
                    
                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).Color = 0

                End If

            Else
                
                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Function Effect_Smoke_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Radius As Integer = 180, Optional ByVal Progression As Single = 1) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Smoke_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Smoke     'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles - Particles * 0.25 * (2 - ClientSetup.bGraphics)       'Set the number of particles
    Effect(EffectIndex).Used = True     'Enabled the effect
    Effect(EffectIndex).X = X          'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx       'Set the graphic
    Effect(EffectIndex).Modifier = Radius       'The direction the effect is animat
    Effect(EffectIndex).Progression = Progression   'Loop the effect

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(15)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).Rhw = 1
        Effect_Smoke_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Smoke_Reset(ByVal EffectIndex As Integer, ByVal index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Reset
'*****************************************************************
    Dim v As Single
    v = Rnd * 20
    'Reset the particle
    Effect(EffectIndex).Particles(index).ResetIt Effect(EffectIndex).X + Effect(EffectIndex).Modifier * RandomNumber(-1, 1) * Rnd / 2, Effect(EffectIndex).Y - Effect(EffectIndex).Modifier, 0, 0, 0, Rnd * -1.5
    Effect(EffectIndex).Particles(index).ResetColor 0.2, 0.2, 0.2, 1, 0

End Sub

Private Sub Effect_Smoke_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime
    
    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime
    
    'Go Through The Particle Loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0.3 Or Effect(EffectIndex).Particles(LoopC).sngY + Effect(EffectIndex).Modifier * 3 < Effect(EffectIndex).Y Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Or Effect(EffectIndex).Progression = -5000 Then
                    
                    
                    'Reset the particle
                    Effect_Smoke_Reset EffectIndex, LoopC
                    
                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).Color = 0

                End If

            Else
                
                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Function Effect_Ray_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Direction As Integer = 180, Optional ByVal Progression As Single = 1) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Ray_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Ray     'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles - Particles * 0.25 * (2 - ClientSetup.bGraphics)       'Set the number of particles
    Effect(EffectIndex).Used = True     'Enabled the effect
    Effect(EffectIndex).X = X          'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx       'Set the graphic
    Effect(EffectIndex).Direction = Direction       'The direction the effect is animat
    Effect(EffectIndex).Progression = Progression   'Loop the effect

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(15)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).Rhw = 1
        Effect_Ray_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Ray_Reset(ByVal EffectIndex As Integer, ByVal index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Reset
'*****************************************************************
    Dim RG As Single
    RG = (Rnd * 1)
    'Reset the particle
    Effect(EffectIndex).Particles(index).ResetIt Effect(EffectIndex).X - 10 + Rnd * 20, Effect(EffectIndex).Y - 10 + Rnd * 20, Rnd * 2 * RandomNumber(-1, 1), Rnd * 2 * RandomNumber(-1, 1), 0, 0
    Effect(EffectIndex).Particles(index).ResetColor RG, RG, 1, 0.8 + (Rnd * 0.2), 0.1 + (Rnd * 4.09)

End Sub

Private Sub Effect_Ray_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime
    
    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime
    
    'Go Through The Particle Loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update The Particle
            'If EffectIndex = EIndex Or EffectIndex = EIndex2 Then
            '    Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime, True
            'Else
                Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime
            'End If
            
            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0.3 And RandomNumber(1, 5) = 1 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Or Effect(EffectIndex).Progression = -5000 Then
                    
                    
                    'Reset the particle
                    Effect_Ray_Reset EffectIndex, LoopC
                    
                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).Color = 0

                End If

            Else
                
                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub
Function Effect_Spell_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional SizeP As Single = 4, Optional Size As Byte = 10, Optional R As Single = 1, Optional G As Single = 1, Optional B As Single = 1, Optional ByVal Direction As Integer = 180, Optional ByVal Progression As Single = 1, Optional Ray As Byte = 5, Optional A As Single = 4.09) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Spell_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Spell     'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles - Particles * 0.25 * (2 - ClientSetup.bGraphics)       'Set the number of particles
    Effect(EffectIndex).Used = True     'Enabled the effect
    Effect(EffectIndex).X = X          'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx       'Set the graphic
    Effect(EffectIndex).Direction = Direction       'The direction the effect is animat
    Effect(EffectIndex).Progression = Progression   'Loop the effect
    
    Effect(EffectIndex).R = R
    Effect(EffectIndex).G = G
    Effect(EffectIndex).B = B
    Effect(EffectIndex).Size = Size
    Effect(EffectIndex).Ray = Ray
    Effect(EffectIndex).SizeP = SizeP
    
    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount
    
    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(SizeP)  'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).Rhw = 1
        Effect_Spell_Reset EffectIndex, LoopC, Size, R, G, B, A
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Spell_Reset(ByVal EffectIndex As Integer, ByVal index As Long, ByVal Size As Byte, ByVal R As Single, ByVal G As Single, ByVal B As Single, A As Single)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Reset
'*****************************************************************

    'Reset the particle
    Effect(EffectIndex).Particles(index).ResetIt Effect(EffectIndex).X - 10 + Rnd * Size, Effect(EffectIndex).Y - 10 + Rnd * Size, Rnd * 0 * RandomNumber(-1, 1), Rnd * 0 * RandomNumber(-1, 1), 0, 0
    Effect(EffectIndex).Particles(index).ResetColor R / 2 + (R / 2 * Rnd), G / 2 + (G / 2 * Rnd), B / 2 + (B / 2 * Rnd), 0.8 + (Rnd * 0.2), 0.1 + (Rnd * A)

End Sub

Private Sub Effect_Spell_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime
    
    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime
    
    'Go Through The Particle Loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0.3 And RandomNumber(1, Effect(EffectIndex).Ray) = 1 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Then 'Or Effect(EffectIndex).Progression = -5000 Then
                    
                    
                    'Reset the particle
                    Effect_Spell_Reset EffectIndex, LoopC, Effect(EffectIndex).Size, Effect(EffectIndex).R, Effect(EffectIndex).G, Effect(EffectIndex).B, Effect(EffectIndex).A
                    
                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).Color = 0

                End If

            Else
                
                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Function Effect_Necro_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Direction As Integer = 180, Optional ByVal Progression As Single = 1) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Necro_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Necro     'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles - Particles * 0.25 * (2 - ClientSetup.bGraphics)       'Set the number of particles
    Effect(EffectIndex).Used = True     'Enabled the effect
    Effect(EffectIndex).X = X          'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx       'Set the graphic
    Effect(EffectIndex).Direction = Direction       'The direction the effect is animat
    Effect(EffectIndex).Progression = Progression   'Loop the effect

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(16)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)
    Effect(EffectIndex).TargetAA = 0
    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).Rhw = 1
        'Effect_Necro_Reset EffectIndex, LoopC
    Next LoopC
    
    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Necro_Reset(ByVal EffectIndex As Integer, ByVal index As Long)

    'Static TargetA As Single
    Dim Co As Single
    Dim Si As Single
    'Calculate the angle
    
    If Effect(EffectIndex).TargetAA = 0 And Effect(EffectIndex).GoToX <> -30000 Then Effect(EffectIndex).TargetAA = Engine_GetAngle(Effect(EffectIndex).X, Effect(EffectIndex).Y, Effect(EffectIndex).GoToX, Effect(EffectIndex).GoToY) + 180
    
    Si = Sin(Effect(EffectIndex).TargetAA * DegreeToRadian)
    Co = Cos(Effect(EffectIndex).TargetAA * DegreeToRadian)
    
    'Reset the particle
    If RandomNumber(1, 2) = 2 Then
        'Effect(EffectIndex).Particles(index).ResetIt Effect(EffectIndex).x, Effect(EffectIndex).y, Co * Sin(Effect(EffectIndex).Progression * 3) * 20, Si * Sin(Effect(EffectIndex).Progression * 3) * 20, 0, 0
        Effect(EffectIndex).Particles(index).ResetIt Effect(EffectIndex).X + Co * Sin(Effect(EffectIndex).Progression) * 25, Effect(EffectIndex).Y + Si * Sin(Effect(EffectIndex).Progression) * 25, 0, 0, 0, 0
        Effect(EffectIndex).Particles(index).ResetColor 0.2, 0.2 + (Rnd * 0.5), 1, 0.5 + (Rnd * 0.2), 0.1 + (Rnd * 4.09)
    Else
        'Effect(EffectIndex).Particles(index).ResetIt Effect(EffectIndex).x, Effect(EffectIndex).y, Co * Sin(Effect(EffectIndex).Progression * 3) * -20, Si * Sin(Effect(EffectIndex).Progression * 3) * -20, 0, 0
        Effect(EffectIndex).Particles(index).ResetIt Effect(EffectIndex).X + Co * Sin(Effect(EffectIndex).Progression) * -25, Effect(EffectIndex).Y + Si * Sin(Effect(EffectIndex).Progression) * -25, 0, 0, 0, 0
        Effect(EffectIndex).Particles(index).ResetColor 1, 0.2 + (Rnd * 0.5), 0.2, 0.7 + (Rnd * 0.2), 0.1 + (Rnd * 4.09)
    End If
    

End Sub

Private Sub Effect_Necro_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime
    
    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime
    
    'Go Through The Particle Loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0.5 And RandomNumber(1, 3) = 3 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Or Effect(EffectIndex).Progression = -5000 Then
                    
                    
                    'Reset the particle
                    Effect_Necro_Reset EffectIndex, LoopC
                    
                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).Color = 0

                End If

            Else
                
                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Function Effect_Curse_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Direction As Integer = 180, Optional ByVal Progression As Single = 1) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Curse_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Curse     'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles - Particles * 0.25 * (2 - ClientSetup.bGraphics)       'Set the number of particles
    Effect(EffectIndex).Used = True     'Enabled the effect
    Effect(EffectIndex).X = X          'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx       'Set the graphic
    Effect(EffectIndex).Direction = Direction       'The direction the effect is animat
    Effect(EffectIndex).Progression = Progression   'Loop the effect

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(12)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)
    Effect(EffectIndex).TargetAA = 0
    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).Rhw = 1
        'Effect_Necro_Reset EffectIndex, LoopC
    Next LoopC
    
    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Curse_Reset(ByVal EffectIndex As Integer, ByVal index As Long)

    'Static TargetA As Single
    Dim Co As Single
    Dim Si As Single
    Dim RG As Single
       
    RG = (Rnd * 0.4)
    'Calculate the angle
    
    If Effect(EffectIndex).TargetAA = 0 And Effect(EffectIndex).GoToX <> -30000 Then Effect(EffectIndex).TargetAA = Engine_GetAngle(Effect(EffectIndex).X, Effect(EffectIndex).Y, Effect(EffectIndex).GoToX, Effect(EffectIndex).GoToY) + 180
    
    Si = Sin(Effect(EffectIndex).TargetAA * DegreeToRadian)
    Co = Cos(Effect(EffectIndex).TargetAA * DegreeToRadian)
    
    'Reset the particle
    If RandomNumber(1, 2) = 2 Then
        'Effect(EffectIndex).Particles(index).ResetIt Effect(EffectIndex).x, Effect(EffectIndex).y, Co * Sin(Effect(EffectIndex).Progression * 3) * 20, Si * Sin(Effect(EffectIndex).Progression * 3) * 20, 0, 0
        Effect(EffectIndex).Particles(index).ResetIt Effect(EffectIndex).X + Co * Sin(Effect(EffectIndex).Progression) * 15, Effect(EffectIndex).Y + Si * Sin(Effect(EffectIndex).Progression) * 15, 0, 0, 0, 0
        Effect(EffectIndex).Particles(index).ResetColor RG, 0.4, RG, 0.5 + (Rnd * 0.2), 0.1 + (Rnd * 4.09)
    Else
        'Effect(EffectIndex).Particles(index).ResetIt Effect(EffectIndex).x, Effect(EffectIndex).y, Co * Sin(Effect(EffectIndex).Progression * 3) * -20, Si * Sin(Effect(EffectIndex).Progression * 3) * -20, 0, 0
        Effect(EffectIndex).Particles(index).ResetIt Effect(EffectIndex).X + Co * Sin(Effect(EffectIndex).Progression) * -15, Effect(EffectIndex).Y + Si * Sin(Effect(EffectIndex).Progression) * -15, 0, 0, 0, 0
        Effect(EffectIndex).Particles(index).ResetColor RG, RG, 0.4, 0.7 + (Rnd * 0.2), 0.1 + (Rnd * 4.09)
    End If
    

End Sub

Private Sub Effect_Curse_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime
    
    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime
    
    'Go Through The Particle Loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0.5 And RandomNumber(1, 3) = 3 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Or Effect(EffectIndex).Progression = -5000 Then
                    
                    
                    'Reset the particle
                    Effect_Curse_Reset EffectIndex, LoopC
                    
                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).Color = 0

                End If

            Else
                
                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub


Function Effect_Ice_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Progression As Single = 1, Optional Looping As Boolean = False) As Integer
'*****************************************************************
'Particle effect template for effects as described on the
'wiki page: http://www.vbgore.com/Particle_effect_equations
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_EquationTemplate_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Ice_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Ice       'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles '- Particles * 0.25 * (2 - ClientSetup.bGraphics)       'Set the number of particles
    Effect(EffectIndex).Used = True                     'Enable the effect
    Effect(EffectIndex).X = X                           'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx                       'Set the graphic
    Effect(EffectIndex).Progression = Progression       'If we loop the effect
    Effect(EffectIndex).Looping = Looping
    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(16)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).Rhw = 1
        Effect_Ice_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Ice_Reset(ByVal EffectIndex As Integer, ByVal index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_EquationTemplate_Reset
'*****************************************************************
Dim X As Single
Dim Y As Single
Dim R As Single
Dim RG As Single
    
    RG = (Rnd * 1)
    
    Effect(EffectIndex).Progression = Effect(EffectIndex).Progression + 0.1
    R = (index / 20) * exp(index / Effect(EffectIndex).Progression Mod 3)
    X = R * Cos(index) * (Rnd * 1.5)
    Y = R * Sin(index) * (Rnd * 1.5)
    
    'Reset the particle
    Effect(EffectIndex).Particles(index).ResetIt Effect(EffectIndex).X + X, Effect(EffectIndex).Y + Y, 0, 0, 0, 0
    Effect(EffectIndex).Particles(index).ResetColor RG, RG, 1, 0.9, 0.2 + (Rnd * 0.2)

End Sub

Private Sub Effect_Ice_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_EquationTemplate_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    
    
    'Go Through The Particle Loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update the particle
            'If EffectIndex = EIndex Or EffectIndex = EIndex2 Then
            '    Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime, True
            'Else
                Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime
            'End If

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0.2 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression < 80 Then

                    'Reset the particle
                    Effect_Ice_Reset EffectIndex, LoopC

                ElseIf Effect(EffectIndex).Looping Then
                    Effect(EffectIndex).Progression = 70
                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).Color = 0

                End If

            Else

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Function Effect_Green_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Progression As Single = 1) As Integer
'*****************************************************************
'Particle effect template for effects as described on the
'wiki page: http://www.vbgore.com/Particle_effect_equations
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_EquationTemplate_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Green_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Green       'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles '- Particles * 0.25 * (2 - ClientSetup.bGraphics)       'Set the number of particles
    Effect(EffectIndex).Used = True                     'Enable the effect
    Effect(EffectIndex).X = X                           'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx                       'Set the graphic
    Effect(EffectIndex).Progression = Progression       'If we loop the effect

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(16)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).Rhw = 1
        Effect_Green_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Green_Reset(ByVal EffectIndex As Integer, ByVal index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_EquationTemplate_Reset
'*****************************************************************
Dim X As Single
Dim Y As Single
Dim R As Single
Dim RG As Single
    
    RG = (Rnd * 0.5)
    
    Effect(EffectIndex).Progression = Effect(EffectIndex).Progression + 0.1
    R = (index / 250) * ((Effect(EffectIndex).Progression / 10) ^ 2)
    X = R * Round(Cos(index), 0) + (index * Rnd * 0.07) * Sgn(Cos(index))
    Y = R * Round(Sin(index), 0) + (index * Rnd * 0.07) * Sgn(Sin(index))
    
    'Reset the particle
    Effect(EffectIndex).Particles(index).ResetIt Effect(EffectIndex).X + X, Effect(EffectIndex).Y + Y, 0, 0, 0, 0
    If RandomNumber(1, 2) = 1 Then
        Effect(EffectIndex).Particles(index).ResetColor 1, 0.2 + RG, 0.2, 0.9, 0.2 + (Rnd * 0.2)
    Else
        Effect(EffectIndex).Particles(index).ResetColor 0.2, 0.2 + RG, 1, 0.9, 0.2 + (Rnd * 0.2)
    End If
    

End Sub

Private Sub Effect_Green_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_EquationTemplate_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    
    
    'Go Through The Particle Loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0.2 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression < 80 Then

                    'Reset the particle
                    Effect_Green_Reset EffectIndex, LoopC

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).Color = 0

                End If

            Else

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Function Effect_Lissajous_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Progression As Single = 1) As Integer

Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Lissajous_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Lissajous       'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True                     'Enable the effect
    Effect(EffectIndex).X = X                           'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx                       'Set the graphic
    Effect(EffectIndex).Progression = Progression       'If we loop the effect

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(8)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).Rhw = 1
        Effect_Lissajous_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Lissajous_Reset(ByVal EffectIndex As Integer, ByVal index As Long)

Dim X As Single
Dim Y As Single

Dim A As Single
Dim B As Integer
Dim Al As Single
Dim RG As Single
    Al = 3.1415 / 2
    A = 3
    B = 4
    
    RG = (Rnd * 0.4)
    
    Effect(EffectIndex).Progression = Effect(EffectIndex).Progression + 0.14
    
    X = A * Sin(Effect(EffectIndex).Progression * Al / 12) * 7 + (Rnd * 10) - 5
    Y = B * Sin(Effect(EffectIndex).Progression / 12) * 7 - 10 + (Rnd * 10) - 5
    
    'Reset the particle
    
    If RandomNumber(1, 2) = 1 Then
        Effect(EffectIndex).Particles(index).ResetIt Effect(EffectIndex).X + X, Effect(EffectIndex).Y + Y, 0, 0, 0, 0
        Effect(EffectIndex).Particles(index).ResetColor RG, 0.4, RG, 0.9, 0.2 + (Rnd * 0.2)
    Else
        Effect(EffectIndex).Particles(index).ResetIt Effect(EffectIndex).X - X, Effect(EffectIndex).Y - Y - 20, 0, 0, 0, 0
        Effect(EffectIndex).Particles(index).ResetColor RG, RG, 0.4, 0.5 + (Rnd * 0.2), 0.2
    End If
    

End Sub

Private Sub Effect_Lissajous_Update(ByVal EffectIndex As Integer)

Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    
    'Effect(EffectIndex).Progression = Effect(EffectIndex).Progression + 0.001
    'Go Through The Particle Loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0.5 And RandomNumber(1, 4) = 1 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression < 200 Then

                    'Reset the particle
                    Effect_Lissajous_Reset EffectIndex, LoopC

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).Color = 0

                End If

            Else

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub



Private Function Effect_FToDW(f As Single) As Long
'*****************************************************************
'Converts a float to a D-Word, or in Visual Basic terms, a Single to a Long
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_FToDW
'*****************************************************************
Dim Buf As D3DXBuffer

    'Converts a single into a long (Float to DWORD)
    Set Buf = D3DX.CreateBuffer(4)
    D3DX.BufferSetData Buf, 0, 4, 1, f
    D3DX.BufferGetData Buf, 0, 4, 1, Effect_FToDW

End Function

Function Effect_Heal_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Progression As Single = 1) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Heal_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Heal_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Heal      'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles - Particles * 0.25 * (2 - ClientSetup.bGraphics)       'Set the number of particles
    Effect(EffectIndex).Used = True     'Enabled the effect
    Effect(EffectIndex).X = X           'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx       'Set the graphic
    Effect(EffectIndex).Progression = Progression   'Loop the effect
    Effect(EffectIndex).KillWhenAtTarget = True     'End the effect when it reaches the target (progression = 0)
    Effect(EffectIndex).KillWhenTargetLost = True   'End the effect if the target is lost (progression = 0)
    
    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(16)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).Rhw = 1
        Effect_Heal_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Heal_Reset(ByVal EffectIndex As Integer, ByVal index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Heal_Reset
'*****************************************************************

    'Reset the particle
    Effect(EffectIndex).Particles(index).ResetIt Effect(EffectIndex).X - 10 + Rnd * 20, Effect(EffectIndex).Y - 10 + Rnd * 20, -Sin((180 + (Rnd * 90) - 45) * 0.0174533) * 8 + (Rnd * 3), Cos((180 + (Rnd * 90) - 45) * 0.0174533) * 8 + (Rnd * 3), 0, 0
    Effect(EffectIndex).Particles(index).ResetColor 0.8, 0.2, 0.2, 0.6 + (Rnd * 0.2), 0.01 + (Rnd * 0.5)
    
End Sub

Private Sub Effect_Heal_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Heal_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim LoopC As Long
Dim i As Integer

    'Calculate the time difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime
    
    'Go through the particle loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression <> 0 Then

                    'Reset the particle
                    Effect_Heal_Reset EffectIndex, LoopC

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).Color = 0

                End If

            Else
                
                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Sub Effect_Kill(ByVal EffectIndex As Integer, Optional ByVal KillAll As Boolean = False)
'*****************************************************************
'Kills (stops) a single effect or all effects
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Kill
'*****************************************************************
Dim LoopC As Long

    'Check If To Kill All Effects
    If KillAll = True Then

        'Loop Through Every Effect
        For LoopC = 1 To NumEffects

            'Stop The Effect
            Effect(LoopC).Used = False

        Next
        
    Else

        'Stop The Selected Effect
        Effect(EffectIndex).Used = False
        
    End If

End Sub

Private Function Effect_NextOpenSlot() As Integer
'*****************************************************************
'Finds the next open effects index
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_NextOpenSlot
'*****************************************************************
Dim EffectIndex As Integer

    'Find The Next Open Effect Slot
    Do
        EffectIndex = EffectIndex + 1   'Check The Next Slot
        If EffectIndex > NumEffects Then    'Dont Go Over Maximum Amount
            Effect_NextOpenSlot = -1
            Exit Function
        End If
    Loop While Effect(EffectIndex).Used = True    'Check Next If Effect Is In Use

    'Return the next open slot
    Effect_NextOpenSlot = EffectIndex

    'Clear the old information from the effect
    Erase Effect(EffectIndex).Particles()
    Erase Effect(EffectIndex).PartVertex()
    ZeroMemory Effect(EffectIndex), LenB(Effect(EffectIndex))
    Effect(EffectIndex).GoToX = -30000
    Effect(EffectIndex).GoToY = -30000

End Function

Function Effect_Protection_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Size As Byte = 30, Optional ByVal Time As Single = 10) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Protection_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Protection_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Protection    'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles - Particles * 0.25 * (2 - ClientSetup.bGraphics)           'Set the number of particles
    Effect(EffectIndex).Used = True             'Enabled the effect
    Effect(EffectIndex).X = X                   'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                   'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx               'Set the graphic
    Effect(EffectIndex).Modifier = Size         'How large the circle is
    Effect(EffectIndex).Progression = Time      'How long the effect will last

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(20)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).Rhw = 1
        Effect_Protection_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Protection_Reset(ByVal EffectIndex As Integer, ByVal index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Protection_Reset
'*****************************************************************
Dim A As Single
Dim X As Single
Dim Y As Single

    'Get the positions
    A = Rnd * 360 * DegreeToRadian
    X = Effect(EffectIndex).X - (Sin(A) * Effect(EffectIndex).Modifier)
    Y = Effect(EffectIndex).Y + (Cos(A) * Effect(EffectIndex).Modifier)

    'Reset the particle
    Effect(EffectIndex).Particles(index).ResetIt X, Y, 0, Rnd * -1, 0, -2
    Effect(EffectIndex).Particles(index).ResetColor 0.1, 0.1, 0.9, 0.6 + (Rnd * 0.4), 0.06 + (Rnd * 0.2)

End Sub

Public Sub Effect_UpdateOffset(ByVal EffectIndex As Integer)
'***************************************************
'Update an effect's position if the screen has moved
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_UpdateOffset
'***************************************************
    If EffectIndex <> CharList(UserCharIndex).AuraIndex Then
        Effect(EffectIndex).X = Effect(EffectIndex).X + (LastOffsetX - ParticleOffsetX) '- Sgn(CharList(UserCharIndex).ScrollDirectionX) * 0.5 * (LastOffsetX - ParticleOffsetX)
        Effect(EffectIndex).Y = Effect(EffectIndex).Y + (LastOffsetY - ParticleOffsetY) '- Sgn(CharList(UserCharIndex).ScrollDirectionY) * 0.5 * (LastOffsetY - ParticleOffsetY)
        'Effect(EffectIndex).x = Effect(EffectIndex).x + ParticleOffsetX
        'Effect(EffectIndex).y = Effect(EffectIndex).y + ParticleOffsetY
    End If

End Sub

Public Sub Effect_UpdateBinding(ByVal EffectIndex As Integer)
 
'***************************************************
'Updates the binding of a particle effect to a target, if
'the effect is bound to a character
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_UpdateBinding
'***************************************************
Dim TargetI As Integer
Dim TargetA As Single
Dim LoopC As Integer
Dim RetNum As Integer

    Effect(EffectIndex).NowMove = timeGetTime
    
    If Effect(EffectIndex).PreviousMove + 10 < Effect(EffectIndex).NowMove Then
        Effect(EffectIndex).PreviousMove = timeGetTime

        'Update position through character binding
        If Effect(EffectIndex).BindToChar > 0 Then
     
            'Store the character index
            TargetI = Effect(EffectIndex).BindToChar
     
            'Check for a valid binding index
            If TargetI > LastChar Then
                Effect(EffectIndex).BindToChar = 0
                If Effect(EffectIndex).KillWhenTargetLost Then
                    Effect(EffectIndex).Progression = 0
                    'Effect(EffectIndex).Used = False
                    Exit Sub
                End If
            ElseIf CharList(TargetI).Active = 0 Then
                Effect(EffectIndex).BindToChar = 0
                If Effect(EffectIndex).KillWhenTargetLost Then
                    Effect(EffectIndex).Progression = 0
                    'Effect(EffectIndex).Used = False
                    Exit Sub
                End If
            Else
     
                'Calculate the X and Y positions
                Effect(EffectIndex).GoToX = Engine_TPtoSPX(CharList(Effect(EffectIndex).BindToChar).Pos.X) + 16
                Effect(EffectIndex).GoToY = Engine_TPtoSPY(CharList(Effect(EffectIndex).BindToChar).Pos.Y) + 16
     
            End If
     
        End If
     
        'Move to the new position if needed
        If Effect(EffectIndex).GoToX > -30000 Or Effect(EffectIndex).GoToY > -30000 Then
            If Effect(EffectIndex).GoToX <> Effect(EffectIndex).X Or Effect(EffectIndex).GoToY <> Effect(EffectIndex).Y Then
     
                'Calculate the angle
                TargetA = Engine_GetAngle(Effect(EffectIndex).X, Effect(EffectIndex).Y, Effect(EffectIndex).GoToX, Effect(EffectIndex).GoToY) + 180
    
                'Update the position of the effect
                Effect(EffectIndex).X = Effect(EffectIndex).X - Sin(TargetA * DegreeToRadian) * Effect(EffectIndex).BindSpeed
                Effect(EffectIndex).Y = Effect(EffectIndex).Y + Cos(TargetA * DegreeToRadian) * Effect(EffectIndex).BindSpeed
                
                'For LoopC = 0 To Effect(EffectIndex).ParticleCount
                '    If Effect(EffectIndex).Particles(LoopC).sngA >= 0 Then
                '     Effect(EffectIndex).Particles(LoopC).sngX = Effect(EffectIndex).Particles(LoopC).sngX - Sin(TargetA * DegreeToRadian) * Effect(EffectIndex).BindSpeed
                '     Effect(EffectIndex).Particles(LoopC).sngY = Effect(EffectIndex).Particles(LoopC).sngY + Cos(TargetA * DegreeToRadian) * Effect(EffectIndex).BindSpeed
                '    End If
                'Next
                
                'Check if the effect is close enough to the target to just stick it at the target
                If Effect(EffectIndex).GoToX > -30000 Then
                    If Abs(Effect(EffectIndex).X - Effect(EffectIndex).GoToX) < 10 Then Effect(EffectIndex).X = Effect(EffectIndex).GoToX
                End If
                If Effect(EffectIndex).GoToY > -30000 Then
                    If Abs(Effect(EffectIndex).Y - Effect(EffectIndex).GoToY) < 10 Then Effect(EffectIndex).Y = Effect(EffectIndex).GoToY
                End If
     
                'Check if the position of the effect is equal to that of the target
                If Effect(EffectIndex).X = Effect(EffectIndex).GoToX Then
                    If Effect(EffectIndex).Y = Effect(EffectIndex).GoToY Then
     
                        'For some effects, if the position is reached, we want to end the effect
                        If Effect(EffectIndex).KillWhenAtTarget Then
                            
                            'Explode on impact
                            If Effect(EffectIndex).Progression <> 0 Then
                            
                                If Effect(EffectIndex).EffectNum = EffectNum_Torch Then
                                    RetNum = Effect_EquationTemplate_Begin(Effect(EffectIndex).X, Effect(EffectIndex).Y, 1, 200, 1)  'Tormenta de fuego
                                ElseIf Effect(EffectIndex).EffectNum = EffectNum_Ray Then
                                    RetNum = Effect_Ice_Begin(Effect(EffectIndex).X, Effect(EffectIndex).Y, 2, 150, 40)  'Descarga electrica
                                ElseIf Effect(EffectIndex).EffectNum = EffectNum_Necro Then
                                    Effect(EffectIndex).TargetAA = 0
                                    RetNum = Effect_Green_Begin(Effect(EffectIndex).X, Effect(EffectIndex).Y, 2, 300, 40)  'Apocalipsis
                                ElseIf Effect(EffectIndex).EffectNum = EffectNum_Curse Then
                                    Effect(EffectIndex).TargetAA = 0
                                    RetNum = Effect_Lissajous_Begin(Effect(EffectIndex).X, Effect(EffectIndex).Y, 1, 250, 1)  'Inmovilizar
                                End If
                            
                                If RetNum > 0 Then
                                    Effect(RetNum).BindToChar = Effect(EffectIndex).BindToChar
                                    Effect(RetNum).BindSpeed = 10
                                End If
                            End If
                            
                            Effect(EffectIndex).BindToChar = 0
                            Effect(EffectIndex).Progression = 0
                            'Effect(EffectIndex).Used = False
                            Effect(EffectIndex).GoToX = Effect(EffectIndex).X
                            Effect(EffectIndex).GoToY = Effect(EffectIndex).Y
                        End If
                        Exit Sub    'The effect is at the right position, don't update
     
                    End If
                End If
     
            End If
        End If
    End If
End Sub

Private Sub Effect_Protection_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Protection_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime

    'Go through the particle loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Then

                    'Reset the particle
                    Effect_Protection_Reset EffectIndex, LoopC

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).Color = 0

                End If

            Else

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Public Sub Effect_Render(ByVal EffectIndex As Integer, Optional ByVal SetRenderStates As Boolean = True)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Render
'*****************************************************************
    Dim i As Integer
    'Check if we have the device
    If D3DDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub

    'Set the render state for the size of the particle
    D3DDevice.SetRenderState D3DRS_POINTSIZE, Effect(EffectIndex).FloatSize
    
    'Set the render state to point blitting
    If SetRenderStates Then D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    
    'Set the last texture to a random number to force the engine to reload the texture
    LastTexture = -65489
    
    'If UserMoving Then
    '    Effect(EffectIndex).x = Effect(EffectIndex).x - Sgn(CharList(UserCharIndex).ScrollDirectionX) * 16
    '    Effect(EffectIndex).y = Effect(EffectIndex).y - Sgn(CharList(UserCharIndex).ScrollDirectionY) * 16
    '    For i = 1 To Effect(EffectIndex).ParticleCount
    '        Effect(EffectIndex).PartVertex(i).x = Effect(EffectIndex).PartVertex(i).x - Sgn(CharList(UserCharIndex).ScrollDirectionX) * 16
    '        Effect(EffectIndex).PartVertex(i).y = Effect(EffectIndex).PartVertex(i).y - Sgn(CharList(UserCharIndex).ScrollDirectionY) * 16
    '    Next
    'End If
    
    'Set the texture
    D3DDevice.SetTexture 0, ParticleTexture(Effect(EffectIndex).Gfx)

    'Draw all the particles at once
    D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, Effect(EffectIndex).ParticleCount, Effect(EffectIndex).PartVertex(0), Len(Effect(EffectIndex).PartVertex(0))

    'Reset the render state back to normal
    If SetRenderStates Then D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    
    'If UserMoving Then
    '    Effect(EffectIndex).x = Effect(EffectIndex).x + Sgn(CharList(UserCharIndex).ScrollDirectionX) * 16
    '    Effect(EffectIndex).y = Effect(EffectIndex).y + Sgn(CharList(UserCharIndex).ScrollDirectionY) * 16
    '    For i = 1 To Effect(EffectIndex).ParticleCount
    '        Effect(EffectIndex).PartVertex(i).x = Effect(EffectIndex).PartVertex(i).x + Sgn(CharList(UserCharIndex).ScrollDirectionX) * 16
    '        Effect(EffectIndex).PartVertex(i).y = Effect(EffectIndex).PartVertex(i).y + Sgn(CharList(UserCharIndex).ScrollDirectionY) * 16
    '    Next
    'End If

End Sub

Function Effect_Snow_Begin(ByVal Gfx As Integer, ByVal Particles As Integer) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Snow_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Snow_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Snow      'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True     'Enabled the effect
    Effect(EffectIndex).Gfx = Gfx       'Set the graphic

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(15)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).Rhw = 1
        Effect_Snow_Reset EffectIndex, LoopC, 1
    Next LoopC

    'Set the initial time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Snow_Reset(ByVal EffectIndex As Integer, ByVal index As Long, Optional ByVal FirstReset As Byte = 0)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Snow_Reset
'*****************************************************************

    If FirstReset = 1 Then

        'The very first reset
        Effect(EffectIndex).Particles(index).ResetIt -200 + (Rnd * (frmMain.renderer.ScaleWidth + 400)), Rnd * (frmMain.renderer.ScaleHeight + 50), Rnd * 5, 5 + Rnd * 3, 0, 0

    Else

        'Any reset after first
        Effect(EffectIndex).Particles(index).ResetIt -200 + (Rnd * (frmMain.renderer.ScaleWidth + 400)), -15 - Rnd * 185, Rnd * 5, 5 + Rnd * 3, 0, 0
        If Effect(EffectIndex).Particles(index).sngX < -20 Then Effect(EffectIndex).Particles(index).sngY = Rnd * (frmMain.renderer.ScaleHeight + 50)
        If Effect(EffectIndex).Particles(index).sngX > frmMain.renderer.ScaleWidth Then Effect(EffectIndex).Particles(index).sngY = Rnd * (frmMain.renderer.ScaleHeight + 50)
        If Effect(EffectIndex).Particles(index).sngY > frmMain.renderer.ScaleHeight Then Effect(EffectIndex).Particles(index).sngX = Rnd * (frmMain.renderer.ScaleWidth + 50)

    End If

    'Set the color
    Effect(EffectIndex).Particles(index).ResetColor 1, 1, 1, 0.8, 0

End Sub

Private Sub Effect_Snow_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Snow_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate the time difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Go through the particle loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check if particle is in use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if to reset the particle
            If Effect(EffectIndex).Particles(LoopC).sngX < -200 Then Effect(EffectIndex).Particles(LoopC).sngA = 0
            If Effect(EffectIndex).Particles(LoopC).sngX > (frmMain.renderer.ScaleWidth + 200) Then Effect(EffectIndex).Particles(LoopC).sngA = 0
            If Effect(EffectIndex).Particles(LoopC).sngY > (frmMain.renderer.ScaleHeight + 200) Then Effect(EffectIndex).Particles(LoopC).sngA = 0

            'Time for a reset, baby!
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0 Then

                'Reset the particle
                Effect_Snow_Reset EffectIndex, LoopC

            Else

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Function Effect_Strengthen_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Size As Byte = 30, Optional ByVal Time As Single = 10, Optional yellow As Boolean = False) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Strengthen_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Strengthen_Begin = EffectIndex

    'Set the effect's variables
    Effect(EffectIndex).EffectNum = EffectNum_Strengthen    'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles - Particles * 0.25 * (2 - ClientSetup.bGraphics)           'Set the number of particles
    Effect(EffectIndex).Used = True             'Enabled the effect
    Effect(EffectIndex).X = X                   'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                   'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx               'Set the graphic
    Effect(EffectIndex).Modifier = Size         'How large the circle is
    Effect(EffectIndex).Progression = Time      'How long the effect will last

    If yellow Then Effect(EffectIndex).R = 5
    
    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(20)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).Rhw = 1
        Effect_Strengthen_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Strengthen_Reset(ByVal EffectIndex As Integer, ByVal index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Strengthen_Reset
'*****************************************************************
Dim A As Single
Dim X As Single
Dim Y As Single

    'Get the positions
    A = Rnd * 360 * DegreeToRadian
    X = Effect(EffectIndex).X - (Sin(A) * Effect(EffectIndex).Modifier)
    Y = Effect(EffectIndex).Y + (Cos(A) * Effect(EffectIndex).Modifier)

    'Reset the particle
    Effect(EffectIndex).Particles(index).ResetIt X, Y, 0, Rnd * -1, 0, -2
    
    If Effect(EffectIndex).R = 5 Then
        Effect(EffectIndex).Particles(index).ResetColor 1, 1, 0.2, 0.6 + (Rnd * 0.4), 0.06 + (Rnd * 0.2)
    Else
        Effect(EffectIndex).Particles(index).ResetColor 0.2, 1, 0.2, 0.6 + (Rnd * 0.4), 0.06 + (Rnd * 0.2)
    End If
    
End Sub

Private Sub Effect_Strengthen_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Strengthen_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate the time difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime

    'Go through the particle loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check if particle is in use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update the particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Then

                    'Reset the particle
                    Effect_Strengthen_Reset EffectIndex, LoopC

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).Color = 0

                End If

            Else

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Sub Effect_UpdateAll()
'*****************************************************************
'Updates all of the effects and renders them
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_UpdateAll
'*****************************************************************
Dim LoopC As Long

    'Make sure we have effects
    If NumEffects = 0 Then Exit Sub

    'Set the render state for the particle effects
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE

    'Update every effect in use
    For LoopC = 1 To NumEffects

        'Make sure the effect is in use
        If Effect(LoopC).Used Then
        
            'Update the effect position if the screen has moved
            If LoopC <> EIndex And LoopC <> EIndex2 Then
                Effect_UpdateOffset LoopC
            End If
            
            'Update the effect position if it is binded
            Effect_UpdateBinding LoopC

            'Find out which effect is selected, then update it
            If Effect(LoopC).EffectNum = EffectNum_Fire Then Effect_Fire_Update LoopC
            'If Effect(loopc).EffectNum = EffectNum_Snow Then Effect_Snow_Update loopc
            If Effect(LoopC).EffectNum = EffectNum_Heal Then Effect_Heal_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Bless Then Effect_Bless_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Explode Then Effect_Explode_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Protection Then Effect_Protection_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Strengthen Then Effect_Strengthen_Update LoopC
            'If Effect(loopc).EffectNum = EffectNum_Rain Then Effect_Rain_Update loopc
            If Effect(LoopC).EffectNum = EffectNum_EquationTemplate Then Effect_EquationTemplate_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Waterfall Then Effect_Waterfall_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Summon Then Effect_Summon_Update LoopC
            
            If Effect(LoopC).EffectNum = EffectNum_Aura Then Effect_Aura_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Aura2 Then Effect_Aura2_Update LoopC
            
            If Effect(LoopC).EffectNum = EffectNum_PortalGroso Then Effect_PortalGroso_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Torch Then Effect_Torch_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Light Then Effect_Light_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Smoke Then Effect_Smoke_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Ray Then Effect_Ray_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Spell Then Effect_Spell_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Ice Then Effect_Ice_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Necro Then Effect_Necro_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Green Then Effect_Green_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Lissajous Then Effect_Lissajous_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Curse Then Effect_Curse_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Misile Then Effect_Misile_Update LoopC
            
            If Effect(LoopC).EffectNum = EffectNum_Teleport Then Effect_Teleport_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Fountain Then Effect_Fountain_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Holy Then Effect_Holy_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Nova Then Effect_Nova_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Implode Then Effect_Implode_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_RedFountain Then Effect_RedFountain_Update LoopC
            'If Effect(LoopC).EffectNum = EffectNum_Medit Then Effect_Medit_Update LoopC

            
            If Effect(LoopC).EffectNum <> EffectNum_Snow And Effect(LoopC).EffectNum <> EffectNum_Rain _
                And Effect(LoopC).EffectNum <> EffectNum_Aura And Effect(LoopC).EffectNum <> EffectNum_Aura2 And Effect(LoopC).EffectNum <> EffectNum_Medit And Effect(LoopC).EffectNum <> EffectNum_MeditMAX Then
                'Render the effect
                Effect_Render LoopC, False
            End If

        End If

    Next
    
    'Set the render state back for normal rendering
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA

End Sub
Sub Effect_UpdateAuras()
Dim LoopC As Long

    'Make sure we have effects
    If NumEffects = 0 Then Exit Sub

    'Set the render state for the particle effects
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE

    'Update every effect in use
    For LoopC = 1 To NumEffects
            
        If Effect(LoopC).Used And (Effect(LoopC).EffectNum = EffectNum_Aura Or Effect(LoopC).EffectNum = EffectNum_Aura2) Then
            'Update the effect position if the screen has moved
            'Effect_UpdateOffset LoopC
            'Update the effect position if it is binded
            'Effect_UpdateBinding LoopC

            'Find out which effect is selected, then update it
            
            'If Effect(LoopC).EffectNum = EffectNum_Aura Then Effect_Aura_Update LoopC
            'If Effect(LoopC).EffectNum = EffectNum_Aura2 Then Effect_Aura2_Update LoopC
         
            
                'Render the effect
                Effect_Render LoopC, False

        End If

    Next
    
    'Set the render state back for normal rendering
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
End Sub
Sub Effect_UpdateW()
'*****************************************************************
'Updates all of the effects and renders them
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_UpdateAll
'*****************************************************************
Dim LoopC As Long

    'Make sure we have effects
    If NumEffects = 0 Then Exit Sub

    'Set the render state for the particle effects
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE

    'Update every effect in use
    For LoopC = 1 To NumEffects

        'Make sure the effect is in use
        If Effect(LoopC).Used Then
            'Update the effect position if the screen has moved
            'Effect_UpdateOffset loopc
            'Update the effect position if it is binded
            'Effect_UpdateBinding loopc

            'Find out which effect is selected, then update it
            
            If Effect(LoopC).EffectNum = EffectNum_Snow Then Effect_Snow_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Rain Then Effect_Rain_Update LoopC
         
            If Effect(LoopC).EffectNum = EffectNum_Snow Or Effect(LoopC).EffectNum = EffectNum_Rain Then
                'Render the effect
                Effect_Render LoopC, False
            End If

        End If

    Next
    
    'Set the render state back for normal rendering
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA

End Sub

Function Effect_Rain_Begin(ByVal Gfx As Integer, ByVal Particles As Integer) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Rain_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Rain_Begin = EffectIndex

    'Set the effect's variables
    Effect(EffectIndex).EffectNum = EffectNum_Rain      'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles - Particles * 0.25 * (2 - ClientSetup.bGraphics)       'Set the number of particles
    Effect(EffectIndex).Used = True     'Enabled the effect
    Effect(EffectIndex).Gfx = Gfx       'Set the graphic

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(10)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).Rhw = 1
        Effect_Rain_Reset EffectIndex, LoopC, 1
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Rain_Reset(ByVal EffectIndex As Integer, ByVal index As Long, Optional ByVal FirstReset As Byte = 0)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Rain_Reset
'*****************************************************************

    If FirstReset = 1 Then

        'The very first reset
        Effect(EffectIndex).Particles(index).ResetIt -200 + (Rnd * (frmMain.renderer.ScaleWidth + 400)), Rnd * (frmMain.renderer.ScaleHeight + 50), Rnd * 5, 25 + Rnd * 12, 0, 0

    Else

        'Any reset after first
        Effect(EffectIndex).Particles(index).ResetIt -200 + (Rnd * 1200), -15 - Rnd * 185, Rnd * 5, 25 + Rnd * 12, 0, 0
        If Effect(EffectIndex).Particles(index).sngX < -20 Then Effect(EffectIndex).Particles(index).sngY = Rnd * (frmMain.renderer.ScaleHeight + 50)
        If Effect(EffectIndex).Particles(index).sngX > frmMain.renderer.ScaleWidth Then Effect(EffectIndex).Particles(index).sngY = Rnd * (frmMain.renderer.ScaleHeight + 50)
        If Effect(EffectIndex).Particles(index).sngY > frmMain.renderer.ScaleHeight Then Effect(EffectIndex).Particles(index).sngX = Rnd * (frmMain.renderer.ScaleWidth + 50)

    End If

    'Set the color
    Effect(EffectIndex).Particles(index).ResetColor 1, 1, 1, 0.4, 0

End Sub

Private Sub Effect_Rain_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Rain_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate the time difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Go through the particle loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check if the particle is in use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update the particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if to reset the particle
            If Effect(EffectIndex).Particles(LoopC).sngX < -200 Then Effect(EffectIndex).Particles(LoopC).sngA = 0
            If Effect(EffectIndex).Particles(LoopC).sngX > (frmMain.renderer.ScaleWidth + 200) Then Effect(EffectIndex).Particles(LoopC).sngA = 0
            If Effect(EffectIndex).Particles(LoopC).sngY > (frmMain.renderer.ScaleHeight + 200) Then Effect(EffectIndex).Particles(LoopC).sngA = 0

            'Time for a reset, baby!
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0 Then

                'Reset the particle
                Effect_Rain_Reset EffectIndex, LoopC

            Else

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Public Sub Effect_Begin(ByVal EffectIndex As Integer, ByVal X As Single, ByVal Y As Single, Optional ByVal Direction As Single = 180, Optional ByVal BindToMap As Boolean = False)
'*****************************************************************
'A very simplistic form of initialization for particle effects
'Should only be used for starting map-based effects
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Begin
'*****************************************************************
Dim RetNum As Byte

    Select Case EffectIndex
        Case EffectNum_Torch
            RetNum = Effect_Light_Begin(Engine_TPtoSPX(X) - 16, Engine_TPtoSPY(Y) + 10, 1, 75, 179, -5000)
            RetNum = Effect_Smoke_Begin(Engine_TPtoSPX(X) - 16, Engine_TPtoSPY(Y) + 10, 1, 10, 20, -5000)
        Case EffectNum_Fountain
            RetNum = Effect_Fountain_Begin(Engine_TPtoSPX(X) - 16, Engine_TPtoSPY(Y) + 8, 1, 160)
        Case EffectNum_RedFountain
            RetNum = Effect_RedFountain_Begin(Engine_TPtoSPX(X) - 18, Engine_TPtoSPY(Y) - 3, 1, 320)
        Case EffectNum_Teleport
            RetNum = Effect_Teleport_Begin(Engine_TPtoSPX(X - 1), Engine_TPtoSPY(Y - 1), 1, 500, 48)
        Case EffectNum_Waterfall
            RetNum = Effect_Waterfall_Begin(Engine_TPtoSPX(X) - 30, Engine_TPtoSPY(Y) - 16, 2, 800)
        Case EffectNum_Holy
            RetNum = Effect_Holy_Begin(Engine_TPtoSPX(X) - 16, Engine_TPtoSPY(Y), 3, 10, 179, -5000)
        Case EffectNum_PortalGroso
            RetNum = Effect_PortalGroso_Begin(Engine_TPtoSPX(X) - 16, Engine_TPtoSPY(Y) - 8, 1, 1000)
    End Select
    
    'Bind the effect to the map if needed
    If BindToMap Then Effect(RetNum).BoundToMap = 1
    
End Sub

Function Effect_Waterfall_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Waterfall_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Waterfall_Begin = EffectIndex

    'Set the effect's variables
    Effect(EffectIndex).EffectNum = EffectNum_Waterfall     'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles - Particles * 0.25 * (2 - ClientSetup.bGraphics)           'Set the number of particles
    Effect(EffectIndex).Used = True             'Enabled the effect
    Effect(EffectIndex).X = X                   'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                   'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx               'Set the graphic

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(20)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).Rhw = 1
        Effect_Waterfall_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Waterfall_Reset(ByVal EffectIndex As Integer, ByVal index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Waterfall_Reset
'*****************************************************************

  
    If Int(Rnd * 10) = 1 Then
        Effect(EffectIndex).Particles(index).ResetIt Effect(EffectIndex).X + (Rnd * 90), Effect(EffectIndex).Y + (Rnd * 120), 0, 2 + (Rnd * 6), 0, 2
    Else
        Effect(EffectIndex).Particles(index).ResetIt Effect(EffectIndex).X + (Rnd * 90), Effect(EffectIndex).Y + (Rnd * 10), 0, 8 + (Rnd * 6), 0, 2
    End If
    Effect(EffectIndex).Particles(index).ResetColor 0.1, 0.1, 0.9, 0.5 + (Rnd * 0.4), 0
    
End Sub

Private Sub Effect_Waterfall_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Waterfall_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime

    'Go through the particle loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
    
        With Effect(EffectIndex).Particles(LoopC)
    
            'Check if the particle is in use
            If .Used Then
    
                'Update The Particle
                .UpdateParticle ElapsedTime

                'Check if the particle is ready to die
                If (.sngY > Effect(EffectIndex).Y + 140) Or (.sngA = 0) Then
    
                    'Reset the particle
                    Effect_Waterfall_Reset EffectIndex, LoopC
    
                Else

                    'Set the particle information on the particle vertex
                    Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(.sngR, .sngG, .sngB, .sngA)
                    Effect(EffectIndex).PartVertex(LoopC).X = .sngX
                    Effect(EffectIndex).PartVertex(LoopC).Y = .sngY
    
                End If
    
            End If
            
        End With

    Next LoopC

End Sub

Function Effect_Summon_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Progression As Single = 0) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Summon_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Summon_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Summon    'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True                     'Enable the effect
    Effect(EffectIndex).X = X                           'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx                       'Set the graphic
    Effect(EffectIndex).Progression = Progression       'If we loop the effect

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(8)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).Used = True
        Effect(EffectIndex).PartVertex(LoopC).Rhw = 1
        Effect_Summon_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Summon_Reset(ByVal EffectIndex As Integer, ByVal index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Summon_Reset
'*****************************************************************
Dim X As Single
Dim Y As Single
Dim R As Single
    
    If Effect(EffectIndex).Progression > 1000 Then
        Effect(EffectIndex).Progression = Effect(EffectIndex).Progression + 1.4
    Else
        Effect(EffectIndex).Progression = Effect(EffectIndex).Progression + 0.5
    End If
    R = (index / 30) * exp(index / Effect(EffectIndex).Progression)
    X = R * Cos(index)
    Y = R * Sin(index)
    
    'Reset the particle
    Effect(EffectIndex).Particles(index).ResetIt Effect(EffectIndex).X + X, Effect(EffectIndex).Y + Y, 0, 0, 0, 0
    Effect(EffectIndex).Particles(index).ResetColor 0, Rnd, 0, 0.9, 0.2 + (Rnd * 0.2)
 
End Sub

Private Sub Effect_Summon_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Summon_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Go Through The Particle Loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).Used Then

            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression < 1800 Then

                    'Reset the particle
                    Effect_Summon_Reset EffectIndex, LoopC

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).Used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).Color = 0

                End If

            Else
            
                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).X = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).Y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub
