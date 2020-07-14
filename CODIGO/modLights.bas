Attribute VB_Name = "modLights"
Option Explicit
Private Type Light
    ACTIVE As Boolean 'Do we ignore this light?
    ID As Integer
    MAP_X As Integer 'Coordinates
    MAP_Y As Integer
    color As Long 'Start colour
    RANGE As Byte
    red As Byte
    green As Byte
    blue As Byte
    Direccion As Byte
End Type

'Light list
Dim light_list() As Light
Dim light_count As Integer
Dim light_last As Integer

Public Function Light_Remove(ByVal light_index As Integer) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    'Make sure it's a legal index
    If Light_Check(light_index) Then
        Light_Destroy light_index
        Light_Remove = True
    End If
End Function

Public Function Light_Color_Value_Get(ByVal light_index As Integer, ByRef color_value As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/28/2003
'
'*****************************************************************
    'Make sure it's a legal index
    If Light_Check(light_index) Then
        color_value = light_list(light_index).color
        Light_Color_Value_Get = True
    End If
End Function

Public Function Light_Create(ByVal MAP_X As Integer, ByVal MAP_Y As Integer, ByVal red As Byte, _
                         ByVal green As Byte, ByVal blue As Byte, _
                        Optional ByVal RANGE As Byte = 1, Optional ByVal Direccion As Byte = 0, Optional ByVal ID As Integer) As Integer
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Returns the light_index if successful, else 0
'Edited by Juan Martín Sotuyo Dodero
'**************************************************************
    If InMapBounds(MAP_X, MAP_Y) Then
        'Make sure there is no light in the given map pos
        'If Map_Light_Get(map_x, map_y) <> 0 Then
        '    Light_Create = 0
        '    Exit Function
        'End If
        Light_Create = Light_Next_Open
        Light_Make Light_Create, MAP_X, MAP_Y, RANGE, ID, red, green, blue, Direccion
    End If
End Function

Public Function Light_Move(ByVal light_index As Integer, ByVal MAP_X As Integer, ByVal MAP_Y As Integer) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Returns true if successful, else false
'**************************************************************
    'Make sure it's a legal CharIndex
    If Light_Check(light_index) Then
        'Make sure it's a legal move
        If InMapBounds(MAP_X, MAP_Y) Then
        
            'Move it
            Light_Erase light_index
            light_list(light_index).MAP_X = MAP_X
            light_list(light_index).MAP_Y = MAP_Y
    
            Light_Move = True
            
        End If
    End If
End Function

Public Function Light_Move_By_Head(ByVal light_index As Integer, ByVal Heading As Byte) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 15/05/2002
'Returns true if successful, else false
'**************************************************************
    Dim MAP_X As Integer
    Dim MAP_Y As Integer
    Dim nX As Integer
    Dim nY As Integer
    Dim AddY As Byte
    Dim AddX As Byte
    'Check for valid heading
    If Heading < 1 Or Heading > 8 Then
        Light_Move_By_Head = False
        Exit Function
    End If

    'Make sure it's a legal CharIndex
    If Light_Check(light_index) Then
    
        MAP_X = light_list(light_index).MAP_X
        MAP_Y = light_list(light_index).MAP_Y


        Select Case Heading
            Case north
                AddY = -1
        
            Case east
                AddX = 1
        
            Case south
                AddY = 1
            
            Case west
                AddX = -1
        End Select
        
        nX = MAP_X + AddX
        nY = MAP_Y + AddY
        
        'Make sure it's a legal move
        If InMapBounds(nX, nY) Then
        
            'Move it
            Light_Erase light_index

            light_list(light_index).MAP_X = nX
            light_list(light_index).MAP_Y = nY
    
            Light_Move_By_Head = True
            
        End If
    End If
End Function

Private Sub Light_Make(ByVal light_index As Integer, ByVal MAP_X As Integer, ByVal MAP_Y As Integer, _
                        ByVal RANGE As Byte, ByVal ID As Integer, ByVal red As Byte, _
                         ByVal green As Byte, ByVal blue As Byte, ByVal Direccion As Byte)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
    'Update array size
    If light_index > light_last Then
        light_last = light_index
        ReDim Preserve light_list(1 To light_last)
    End If
    light_count = light_count + 1
    
    'Make active
    light_list(light_index).ACTIVE = True
    
    light_list(light_index).MAP_X = MAP_X
    light_list(light_index).MAP_Y = MAP_Y
    light_list(light_index).red = red
    light_list(light_index).green = green
    light_list(light_index).blue = blue
    light_list(light_index).RANGE = RANGE
    light_list(light_index).Direccion = Direccion
    light_list(light_index).ID = ID

End Sub

Private Function Light_Check(ByVal light_index As Integer) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check light_index
    If light_index > 0 And light_index <= light_last Then
        If light_list(light_index).ACTIVE Then
            Light_Check = True
        End If
    End If
End Function

Public Sub Light_Render_Area()

'   Author: Dunkan
'   Note: Las luces redondas son pesadisimas >.< mejor renderizar solo el area.
'   OPTIMIZAR SUB CUANDO SE PUEDA!!!!!

    Dim i As Integer
            
    For i = 1 To light_count
        If light_list(i).MAP_X > UserPos.X - TileBufferSize - 5 And light_list(i).MAP_X < UserPos.X + TileBufferSize + 5 Then
            If light_list(i).MAP_Y > UserPos.Y - TileBufferSize - 5 And light_list(i).MAP_Y < UserPos.Y + TileBufferSize + 5 Then
                If Light_Check(i) Then Light_Render i
            End If
        End If
    
    Next i

End Sub

Public Function Light_IdbyPos(ByVal X As Integer, ByVal Y As Integer) As Integer

'   Author: Dunkan
'   Note: Las luces redondas son pesadisimas >.< mejor renderizar solo el area.
'   OPTIMIZAR SUB CUANDO SE PUEDA!!!!!

    Dim i As Integer
            
    For i = 1 To light_count
        If light_list(i).MAP_X = X And light_list(i).MAP_Y = Y Then
            Light_IdbyPos = i
            Exit For
        End If
    Next i

End Function


Public Sub Light_Render_All()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Modify By Dunkan
'**************************************************************
   
    Dim loop_counter As Integer
            
    For loop_counter = 1 To light_count
        'If Client_Setup.Effect_LightType = 1 Then '1 = Redondas
            If Light_Check(loop_counter) Then Light_Render loop_counter
        'Else '0 = Cuadradas FEAAAA xD
        '    If Light_Check(loop_counter) Then Map_LightRender_Square loop_counter
        'End If
    Next loop_counter
    
End Sub

Private Function CalcularRadio(ByVal cRadio As Integer, ByVal LightX As Integer, ByVal LightY As Integer, ByVal XCoordenadas As Integer, ByVal YCoordenadas As Integer, TileLight As Long, LightColor As D3DCOLORVALUE, AmbientColor As D3DCOLORVALUE) As Long
    Dim DistanciaX As Single
    Dim DistanciaY As Single
    Dim DistanciaVertex As Single
    Dim Radio As Integer
    
    Dim CurrentColor As D3DCOLORVALUE
    
    Radio = cRadio
    
    DistanciaX = LightX + 0.5 - XCoordenadas
    DistanciaY = LightY + 0.5 - YCoordenadas
    
    DistanciaVertex = Sqr(DistanciaX * DistanciaX + DistanciaY * DistanciaY)
    
    If DistanciaVertex <= Radio Then
        Call D3DXColorLerp(CurrentColor, LightColor, AmbientColor, DistanciaVertex / Radio)
        CalcularRadio = D3DColorXRGB(CurrentColor.R, CurrentColor.G, CurrentColor.B)
        If TileLight > CalcularRadio Then CalcularRadio = TileLight
    Else
        CalcularRadio = TileLight
    End If
End Function

Private Sub Light_Render(ByVal light_index As Integer)

    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim ia As Single
    Dim i As Integer
    Dim color As Long
    Dim Ya As Integer
    Dim Xa As Integer
    Dim TileLight As D3DCOLORVALUE
    Dim LightColor As D3DCOLORVALUE
    
    Dim XCoord As Integer
    Dim YCoord As Integer
    
    LightColor.a = 255
    LightColor.R = light_list(light_index).red
    LightColor.G = light_list(light_index).green
    LightColor.B = light_list(light_index).blue
    
    'Set up light borders
        min_x = light_list(light_index).MAP_X - light_list(light_index).RANGE * IIf(light_list(light_index).Direccion And 1, 0, 1)
        min_y = light_list(light_index).MAP_Y - light_list(light_index).RANGE * IIf(light_list(light_index).Direccion And 2, 0, 1)
        max_x = light_list(light_index).MAP_X + light_list(light_index).RANGE * IIf(light_list(light_index).Direccion And 4, 0, 1)
        max_y = light_list(light_index).MAP_Y + light_list(light_index).RANGE * IIf(light_list(light_index).Direccion And 8, 0, 1)

    
            For Ya = min_y To max_y
            For Xa = min_x To max_x
                If InMapBounds(Xa, Ya) Then
                    XCoord = Xa
                    YCoord = Ya
                    MapData(Xa, Ya).Light_Value(0) = CalcularRadio(light_list(light_index).RANGE, _
                    light_list(light_index).MAP_X, light_list(light_index).MAP_Y, XCoord, _
                    YCoord, MapData(Xa, Ya).Light_Value(0), LightColor, IluRGB)

                    XCoord = Xa + 1
                    YCoord = Ya
                    MapData(Xa, Ya).Light_Value(1) = CalcularRadio(light_list(light_index).RANGE, _
                    light_list(light_index).MAP_X, light_list(light_index).MAP_Y, XCoord, _
                    YCoord, MapData(Xa, Ya).Light_Value(1), LightColor, IluRGB)
                       
                    XCoord = Xa
                    YCoord = Ya + 1
                    MapData(Xa, Ya).Light_Value(2) = CalcularRadio(light_list(light_index).RANGE, _
                    light_list(light_index).MAP_X, light_list(light_index).MAP_Y, XCoord, _
                    YCoord, MapData(Xa, Ya).Light_Value(2), LightColor, IluRGB)
                    
                    XCoord = Xa + 1
                    YCoord = Ya + 1
                    MapData(Xa, Ya).Light_Value(3) = CalcularRadio(light_list(light_index).RANGE, _
                    light_list(light_index).MAP_X, light_list(light_index).MAP_Y, XCoord, _
                    YCoord, MapData(Xa, Ya).Light_Value(3), LightColor, IluRGB)
                End If
            Next Xa
            
        Next Ya

End Sub

Private Function Light_Next_Open() As Integer
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Integer
    
    loopc = 1
    Do Until light_list(loopc).ACTIVE = False
        If loopc = light_last Then
            Light_Next_Open = light_last + 1
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Light_Next_Open = loopc
Exit Function
ErrorHandler:
    Light_Next_Open = 1
End Function

Public Function Light_Find(ByVal ID As Integer) As Integer
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'Find the index related to the handle
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Integer
    
    loopc = 1
    Do Until light_list(loopc).ID = ID
        If loopc = light_last Then
            Light_Find = 0
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Light_Find = loopc
Exit Function
ErrorHandler:
    Light_Find = 0
End Function

Public Function Light_Remove_All() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    Dim Index As Integer
    
    For Index = 1 To light_last
        'Make sure it's a legal index
        If Light_Check(Index) Then
         light_list(Index).red = 150
         light_list(Index).blue = 150
         light_list(Index).green = 150
            Light_Destroy Index
        End If
    Next Index
    
    Light_Remove_All = True
End Function
Public Sub Light_Destroy_ToMap(ByVal X As Integer, ByVal Y As Integer)
    Dim Index As Integer
    
    For Index = 1 To light_last
        If light_list(Index).MAP_X = X And light_list(Index).MAP_Y = Y Then
           light_list(Index).ACTIVE = False
           Light_Destroy Index
           ' Call Light_Remove(Index)
           Exit For
        End If
    Next Index
End Sub
Private Sub Light_Destroy(ByVal light_index As Integer)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim temp As Light
    
    Light_Erase light_index
    
    light_list(light_index) = temp
    
    'Update array size
    If light_index = light_last Then
        Do Until light_list(light_last).ACTIVE
            light_last = light_last - 1
            If light_last = 0 Then
                light_count = 0
                Exit Sub
            End If
        Loop
        ReDim Preserve light_list(1 To light_last)
    End If
    light_count = light_count - 1
End Sub

Private Sub Light_Erase(ByVal light_index As Integer)
'***************************************'
'Author: Juan Martín Sotuyo Dodero
'Last modified: 3/31/2003
'Correctly erases a light
'***************************************'
    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim colorz As Long
    colorz = D3DColorXRGB(IluRGB.R, IluRGB.G, IluRGB.B)
    'Set up light borders
    min_x = light_list(light_index).MAP_X - light_list(light_index).RANGE
    min_y = light_list(light_index).MAP_Y - light_list(light_index).RANGE
    max_x = light_list(light_index).MAP_X + light_list(light_index).RANGE
    max_y = light_list(light_index).MAP_Y + light_list(light_index).RANGE
    
    'Arrange corners
    'NE
    If InMapBounds(min_x, min_y) Then
        MapData(min_x, min_y).Light_Value(2) = colorz
    End If
    'NW
    If InMapBounds(max_x, min_y) Then
        MapData(max_x, min_y).Light_Value(0) = colorz
    End If
    'SW
    If InMapBounds(max_x, max_y) Then
        MapData(max_x, max_y).Light_Value(1) = colorz
    End If
    'SE
    If InMapBounds(min_x, max_y) Then
        MapData(min_x, max_y).Light_Value(3) = colorz
    End If
    
    'Arrange borders
    'Upper border
    For X = min_x + 1 To max_x - 1
        If InMapBounds(X, min_y) Then
            MapData(X, min_y).Light_Value(0) = colorz
            MapData(X, min_y).Light_Value(2) = colorz
        End If
    Next X
    
    'Lower border
    For X = min_x + 1 To max_x - 1
        If InMapBounds(X, max_y) Then
            MapData(X, max_y).Light_Value(1) = colorz
            MapData(X, max_y).Light_Value(3) = colorz
        End If
    Next X
    
    'Left border
    For Y = min_y + 1 To max_y - 1
        If InMapBounds(min_x, Y) Then
            MapData(min_x, Y).Light_Value(2) = colorz
            MapData(min_x, Y).Light_Value(3) = colorz
        End If
    Next Y
    
    'Right border
    For Y = min_y + 1 To max_y - 1
        If InMapBounds(max_x, Y) Then
            MapData(max_x, Y).Light_Value(0) = colorz
            MapData(max_x, Y).Light_Value(1) = colorz
        End If
    Next Y
    
    'Set the inner part of the light
    For X = min_x + 1 To max_x - 1
        For Y = min_y + 1 To max_y - 1
            If InMapBounds(X, Y) Then
                MapData(X, Y).Light_Value(0) = colorz
                MapData(X, Y).Light_Value(1) = colorz
                MapData(X, Y).Light_Value(2) = colorz
                MapData(X, Y).Light_Value(3) = colorz
            End If
        Next Y
    Next X
    
End Sub
