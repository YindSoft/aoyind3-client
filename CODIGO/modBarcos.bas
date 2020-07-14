Attribute VB_Name = "modBarcos"
Option Explicit

Public Const NUM_PUERTOS As Byte = 5

Public Const PUERTO_NIX As Byte = 1
Public Const PUERTO_BANDER As Byte = 2
Public Const PUERTO_ARGHAL As Byte = 3
Public Const PUERTO_LINDOS As Byte = 4
Public Const PUERTO_ARKHEIN As Byte = 5


Public Type tPuerto
    Paso(0 To 1) As Byte
    nombre As String
End Type

Public Puertos(1 To NUM_PUERTOS) As tPuerto

Public Barco(0 To 1) As clsBarco

Public RutaBarco(0 To 1) As String

Public Sub InitBarcos()
Puertos(PUERTO_NIX).nombre = "Nix"
Puertos(PUERTO_NIX).Paso(0) = 0
Puertos(PUERTO_NIX).Paso(1) = 21

Puertos(PUERTO_BANDER).nombre = "Banderbill"
Puertos(PUERTO_BANDER).Paso(0) = 4
Puertos(PUERTO_BANDER).Paso(1) = 16

Puertos(PUERTO_ARGHAL).nombre = "Arghal"
Puertos(PUERTO_ARGHAL).Paso(0) = 12
Puertos(PUERTO_ARGHAL).Paso(1) = 9

Puertos(PUERTO_LINDOS).nombre = "Lindos"
Puertos(PUERTO_LINDOS).Paso(0) = 15
Puertos(PUERTO_LINDOS).Paso(1) = 5

Puertos(PUERTO_ARKHEIN).nombre = "Arkhein"
Puertos(PUERTO_ARKHEIN).Paso(0) = 19
Puertos(PUERTO_ARKHEIN).Paso(1) = 0

RutaBarco(0) = "161,1247;35,1247;35,22;302,22;302,55;303,55;566,55;566,65;635,65;635,54;800,54;800,307;801,307;870,307;870,999;887,999;870,999;870,1224;643,1224;643,1371;643,1472;195,1472;195,1266;169,1266;169,1247"
RutaBarco(1) = "639,1383;647,1383;647,1228;874,1228;874,995;887,995;874,995;874,303;804,303;804,311;804,50;631,50;631,61;570,61;570,51;306,51;306,59;306,18;31,18;31,1251;165,1251;165,1243;165,1270;191,1270;191,1476;647,1476;647,1383"

End Sub

Public Sub RenderBarcos(ByVal X As Integer, ByVal Y As Integer, ByVal TileX As Integer, ByVal TileY As Integer, ByVal PixelOffSetX As Single, ByVal PixelOffSetY As Single)
Dim i As Byte
For i = 0 To 1
    If Not Barco(i) Is Nothing Then
        If Barco(i).X = X And Barco(i).Y = Y Then
            Call Barco(i).Render(TileX, TileY, PixelOffSetX, PixelOffSetY)
        End If
    End If
Next i
End Sub

Public Sub CalcularBarcos()
Dim i As Byte
For i = 0 To 1
    If Not Barco(i) Is Nothing Then
        Call Barco(i).Calcular
    End If
Next i
End Sub
