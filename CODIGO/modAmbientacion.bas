Attribute VB_Name = "modAmbientacion"
Public Enum TipoPaso
    CONST_BOSQUE = 1
    CONST_NIEVE = 2
    CONST_CABALLO = 3
    CONST_DUNGEON = 4
    CONST_PISO = 5
    CONST_DESIERTO = 6
    CONST_PESADO = 7
End Enum

Public Type tPaso
    CantPasos As Byte
    Wav() As Integer
End Type

Public Const NUM_PASOS As Byte = 7
Public Pasos() As tPaso


Public luz_dia(0 To 24) As D3DCOLORVALUE
Public Iluminacion As Long
Public IluRGB As D3DCOLORVALUE
Public Hora As Byte

Private Function GetTerrenoDePaso(ByVal TerrainFileNum As Integer, Terrain2FileNum) As TipoPaso

If (TerrainFileNum >= 6000 And TerrainFileNum <= 6004) Or (TerrainFileNum >= 550 And TerrainFileNum <= 552) Or (TerrainFileNum >= 6018 And TerrainFileNum <= 6020) Or _
   (TerrainFileNum >= 1478 And TerrainFileNum <= 1487) Or (TerrainFileNum >= 1548 And TerrainFileNum <= 1551) Or (TerrainFileNum >= 10013 And TerrainFileNum <= 10015) Or _
   (TerrainFileNum >= 1073 And TerrainFileNum <= 1074) Or TerrainFileNum = 14638 Or TerrainFileNum = 14656 Or Terrain2FileNum = 8007 Then
    GetTerrenoDePaso = CONST_BOSQUE
    Exit Function
ElseIf (TerrainFileNum >= 7501 And TerrainFileNum <= 7507) Or (TerrainFileNum = 7500 Or TerrainFileNum = 7508 Or TerrainFileNum = 1533 Or TerrainFileNum = 2508) Then
    GetTerrenoDePaso = CONST_DUNGEON
    Exit Function
ElseIf (TerrainFileNum >= 13106 And TerrainFileNum <= 13115) Or Terrain2FileNum = 13117 Then
    GetTerrenoDePaso = CONST_NIEVE
    Exit Function
ElseIf (TerrainFileNum >= 6018 And TerrainFileNum <= 6021) Or (TerrainFileNum >= 14551 And TerrainFileNum <= 14553) Or TerrainFileNum = 14564 Then
    GetTerrenoDePaso = CONST_DESIERTO
    Exit Function
Else
    GetTerrenoDePaso = CONST_PISO
End If

End Function
Public Sub CargarPasos()

ReDim Pasos(1 To NUM_PASOS) As tPaso

Pasos(TipoPaso.CONST_BOSQUE).CantPasos = 2
ReDim Pasos(TipoPaso.CONST_BOSQUE).Wav(1 To Pasos(TipoPaso.CONST_BOSQUE).CantPasos) As Integer
Pasos(TipoPaso.CONST_BOSQUE).Wav(1) = 193
Pasos(TipoPaso.CONST_BOSQUE).Wav(2) = 194

Pasos(TipoPaso.CONST_NIEVE).CantPasos = 2
ReDim Pasos(TipoPaso.CONST_NIEVE).Wav(1 To Pasos(TipoPaso.CONST_NIEVE).CantPasos) As Integer
Pasos(TipoPaso.CONST_NIEVE).Wav(1) = 195
Pasos(TipoPaso.CONST_NIEVE).Wav(2) = 196

Pasos(TipoPaso.CONST_DUNGEON).CantPasos = 2
ReDim Pasos(TipoPaso.CONST_DUNGEON).Wav(1 To Pasos(TipoPaso.CONST_DUNGEON).CantPasos) As Integer
Pasos(TipoPaso.CONST_DUNGEON).Wav(1) = 23
Pasos(TipoPaso.CONST_DUNGEON).Wav(2) = 24

Pasos(TipoPaso.CONST_DESIERTO).CantPasos = 2
ReDim Pasos(TipoPaso.CONST_DESIERTO).Wav(1 To Pasos(TipoPaso.CONST_DESIERTO).CantPasos) As Integer
Pasos(TipoPaso.CONST_DESIERTO).Wav(1) = 197
Pasos(TipoPaso.CONST_DESIERTO).Wav(2) = 198

Pasos(TipoPaso.CONST_PISO).CantPasos = 2
ReDim Pasos(TipoPaso.CONST_PISO).Wav(1 To Pasos(TipoPaso.CONST_PISO).CantPasos) As Integer
Pasos(TipoPaso.CONST_PISO).Wav(1) = 23
Pasos(TipoPaso.CONST_PISO).Wav(2) = 24

Pasos(TipoPaso.CONST_PESADO).CantPasos = 3
ReDim Pasos(TipoPaso.CONST_PESADO).Wav(1 To Pasos(TipoPaso.CONST_PESADO).CantPasos) As Integer
Pasos(TipoPaso.CONST_PESADO).Wav(1) = 220
Pasos(TipoPaso.CONST_PESADO).Wav(2) = 221
Pasos(TipoPaso.CONST_PESADO).Wav(3) = 222

End Sub

Public Sub DoPasosFx(ByVal CharIndex As Integer)
Dim FileNum As Integer
Dim FileNum2 As Integer
Dim TerrenoDePaso As TipoPaso
    If UserNavegando Or HayAgua(charlist(CharIndex).Pos.x, charlist(CharIndex).Pos.y) Then
        Call Audio.Sound_Play(SND_NAVEGANDO, charlist(CharIndex).Pos.x, charlist(CharIndex).Pos.y)
    Else
        With charlist(CharIndex)
            If Not .muerto And EstaPCarea(CharIndex) And (.priv = 0 Or .priv > 5 Or CharIndex = UserCharIndex) Then
            
                FileNum = MapData(.Pos.x, .Pos.y).Graphic(1).GrhIndex
                If FileNum > 0 Then FileNum = GrhData(FileNum).FileNum
                FileNum2 = MapData(.Pos.x, .Pos.y).Graphic(2).GrhIndex
                If FileNum2 > 0 Then FileNum2 = GrhData(FileNum2).FileNum
                    
                TerrenoDePaso = GetTerrenoDePaso(FileNum, FileNum2)
            
            
                .pie = Not .pie
                
                If .pie Then
                    Call Audio.Sound_Play(Pasos(TerrenoDePaso).Wav(1), .Pos.x, .Pos.y)
                Else
                    Call Audio.Sound_Play(Pasos(TerrenoDePaso).Wav(2), .Pos.x, .Pos.y)
                End If
            End If
        End With
    End If
End Sub


Public Sub setup_ambient()

'Noche 87, 61, 43
luz_dia(0).r = 150
luz_dia(0).G = 150
luz_dia(0).b = 150
luz_dia(1).r = 150
luz_dia(1).G = 150
luz_dia(1).b = 150
luz_dia(2).r = 150
luz_dia(2).G = 150
luz_dia(2).b = 150
luz_dia(3).r = 160
luz_dia(3).G = 160
luz_dia(3).b = 160
'4 am 124,117,91
luz_dia(4).r = 170
luz_dia(4).G = 170
luz_dia(4).b = 170
'5,6 am 143,137,135
luz_dia(5).r = 190
luz_dia(5).G = 190
luz_dia(5).b = 190
luz_dia(6).r = 220
luz_dia(6).G = 220
luz_dia(6).b = 220
'7 am 212,205,207
luz_dia(7).r = 230
luz_dia(7).G = 222
luz_dia(7).b = 222
luz_dia(8).r = 235
luz_dia(8).G = 230
luz_dia(8).b = 230
luz_dia(9).r = 240
luz_dia(9).G = 240
luz_dia(9).b = 240
luz_dia(10).r = 250
luz_dia(10).G = 250
luz_dia(10).b = 250
luz_dia(11).r = 250
luz_dia(11).G = 250
luz_dia(11).b = 250
luz_dia(12).r = 255
luz_dia(12).G = 255
luz_dia(12).b = 255
'Dia 255, 255, 255
luz_dia(12).r = 255
luz_dia(12).G = 255
luz_dia(12).b = 255
luz_dia(13).r = 255
luz_dia(13).G = 255
luz_dia(13).b = 255
'Medio Dia 255, 200, 255
luz_dia(14).r = 255
luz_dia(14).G = 250
luz_dia(14).b = 255
luz_dia(15).r = 255
luz_dia(15).G = 240
luz_dia(15).b = 255
luz_dia(16).r = 255
luz_dia(16).G = 230
luz_dia(16).b = 255
'17/18 0, 100, 255
luz_dia(17).r = 245
luz_dia(17).G = 235
luz_dia(17).b = 235
'18/19 0, 100, 255
luz_dia(18).r = 235
luz_dia(18).G = 230
luz_dia(18).b = 235
'19/20 156, 142, 83
luz_dia(19).r = 220
luz_dia(19).G = 210
luz_dia(19).b = 210
luz_dia(20).r = 200
luz_dia(20).G = 180
luz_dia(20).b = 180
luz_dia(21).r = 160
luz_dia(21).G = 160
luz_dia(21).b = 160
luz_dia(22).r = 150
luz_dia(22).G = 150
luz_dia(22).b = 150
luz_dia(23).r = 150
luz_dia(23).G = 150
luz_dia(23).b = 150
luz_dia(24).r = 150
luz_dia(24).G = 150
luz_dia(24).b = 150
End Sub
Public Sub SetDayLight()
Dim pHora As Byte
If Zonas(ZonaActual).Terreno = eTerreno.Dungeon Then
    pHora = 24
Else
    pHora = Hora
End If
IluRGB.r = luz_dia(pHora).r
IluRGB.G = luz_dia(pHora).G
IluRGB.b = luz_dia(pHora).b

Iluminacion = D3DColorRGBA(IluRGB.r, IluRGB.G, IluRGB.b, 255)
ColorTecho = D3DColorRGBA(IluRGB.r, IluRGB.G, IluRGB.b, bAlpha)

End Sub

