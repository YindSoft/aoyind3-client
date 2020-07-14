VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNPCMuerto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Body As BodyData
Private Head As HeadData
Private Arma As WeaponAnimData
Private Escudo As ShieldAnimData
Private Casco As HeadData
Private Heading As E_Heading
Private Pos As Position
Private cErase As Single
Private iTick As Long
Public KillMe As Boolean

Friend Sub Init(pBody As Integer, pHead As Integer, pArma As Integer, pEscudo As Integer, pCasco As Integer, pHeading As E_Heading, pPos As Position, pAlpha As Byte)

Body = BodyData(pBody)
Head = HeadData(pHead)
Arma = WeaponAnimData(pArma)
Escudo = ShieldAnimData(pEscudo)
Casco = CascoAnimData(pCasco)
Heading = pHeading
Pos.x = pPos.x
Pos.y = pPos.y
cErase = pAlpha

End Sub
Public Sub Update()
If cErase > 0 Then
    If iTick < (GetTickCount() And &H7FFFFFFF) - 35 Then
        cErase = cErase - 10
        If cErase <= 0 Then
            Set MapData(Pos.x, Pos.y).Elemento = Nothing
            KillMe = True
            Exit Sub
        Else
            iTick = (GetTickCount() And &H7FFFFFFF)
        End If
    End If
End If
End Sub
Public Sub Render(ByVal PixelOffSetX As Single, ByVal PixelOffSetY As Single)
Dim ColorPj As Long

'PixelOffSetX = (Pos.X - X + 8) * 32 + PixelOffSetX + 256
'PixelOffSetY = (Pos.Y - Y + 6) * 32 + PixelOffSetY + 256

ColorPj = D3DColorRGBA(IluRGB.R, IluRGB.G, IluRGB.B, cErase)
'Draw Body
If Body.Walk(Heading).GrhIndex Then _
    Call DrawGrh(Body.Walk(Heading), PixelOffSetX, PixelOffSetY, 1, 1, 1, ColorPj, cErase)
            
'Draw Head
If Head.Head(Heading).GrhIndex Then _
    Call DrawGrh(Head.Head(Heading), PixelOffSetX + Body.HeadOffset.x, PixelOffSetY + Body.HeadOffset.y, 1, 0, 2, ColorPj, cErase)
    
'Draw Helmet
If Casco.Head(Heading).GrhIndex Then _
    Call DrawGrh(Casco.Head(Heading), PixelOffSetX + Body.HeadOffset.x, PixelOffSetY + Body.HeadOffset.y, 1, 0, 2, ColorPj, cErase)
    
'Draw Weapon
If Arma.WeaponWalk(Heading).GrhIndex Then _
    Call DrawGrh(Arma.WeaponWalk(Heading), PixelOffSetX, PixelOffSetY, 1, 1, 1, ColorPj, cErase)
    
'Draw Shield
If Escudo.ShieldWalk(Heading).GrhIndex Then _
    Call DrawGrh(Escudo.ShieldWalk(Heading), PixelOffSetX, PixelOffSetY, 1, 1, 1, ColorPj, cErase)

End Sub

