Attribute VB_Name = "ModAreas"
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

'LAS GUARDAMOS PARA PROCESAR LOS MPs y sabes si borrar personajes
Public Const MargenX As Integer = 12
Public Const MargenY As Integer = 10

Public Sub CambioDeArea(ByVal X As Integer, ByVal Y As Integer, ByVal Head As Byte)
    Dim loopX As Long, loopY As Long
    Dim MinX As Integer
    Dim MinY As Integer
    Dim MaxX As Integer
    Dim MaxY As Integer
    MinX = X
    MinY = Y
    MaxX = X
    MaxY = Y
        If Head = E_Heading.SOUTH Then
            MinX = MinX - MargenX
            MaxX = MaxX + MargenX
            MinY = MinY - MargenY - 1
            MaxY = MinY
        ElseIf Head = E_Heading.NORTH Then
            MinX = MinX - MargenX
            MaxX = MaxX + MargenX
            MinY = MinY + MargenY + 1
            MaxY = MinY
        
        ElseIf Head = E_Heading.EAST Then
            MinX = MinX - MargenX - 1
            MaxX = MinX
            MinY = MinY - MargenY
            MaxY = MaxY + MargenY
        
        
        ElseIf Head = E_Heading.WEST Then
            MinX = MinX + MargenX + 1
            MaxX = MinX
            MinY = MinY - MargenY
            MaxY = MaxY + MargenY
    
        End If
    
    If MinY < 1 Then MinY = 1
    If MinX < 1 Then MinX = 1
    If MaxY > YMaxMapSize Then MaxY = YMaxMapSize
    If MaxX > XMaxMapSize Then MaxX = XMaxMapSize

    For loopX = MinX To MaxX
        For loopY = MinY To MaxY
                If MapData(loopX, loopY).CharIndex > 0 Then
                    If MapData(loopX, loopY).CharIndex <> UserCharIndex Then
                        Call EraseChar(MapData(loopX, loopY).CharIndex)
                    End If
                End If
                'Erase OBJs
                If MapData(loopX, loopY).ObjGrh.GrhIndex = GrhFogata Then
                    MapData(loopX, loopY).Graphic(3).GrhIndex = 0
                    Call Light_Destroy_ToMap(loopX, loopY)
                End If
                MapData(loopX, loopY).ObjGrh.GrhIndex = 0
        Next loopY
    Next loopX
    
    'Call RefreshAllChars
End Sub

Public Sub LimpiarArea()
Dim X As Integer
Dim Y As Integer
    For X = UserPos.X - MargenX * 2 To UserPos.X + MargenX * 2
        For Y = UserPos.Y - MargenY * 2 To UserPos.Y + MargenY * 2
            If InMapBounds(X, Y) Then
                If MapData(X, Y).CharIndex > 0 Then
                    If MapData(X, Y).CharIndex <> UserCharIndex Then
                        Call EraseChar(MapData(X, Y).CharIndex)
                    End If
                End If
                If MapData(X, Y).ObjGrh.GrhIndex = GrhFogata Then
                    MapData(X, Y).Graphic(3).GrhIndex = 0
                    Call Light_Destroy_ToMap(X, Y)
                End If
                MapData(X, Y).ObjGrh.GrhIndex = 0
            End If
        Next Y
    Next X
End Sub
