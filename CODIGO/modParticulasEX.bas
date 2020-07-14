Attribute VB_Name = "modParticulas"
Option Explicit

Public Const D3DFVF_TLVERTEX = (D3DFVF_XYZRHW Or D3DFVF_DIFFUSE Or D3DFVF_TEX1)
    
Public Type typeTRANSLITVERTEX
    X As Single
    Y As Single
    z As Single
    rhw As Single
    color As Long
    tu As Single
    tv As Single
End Type

Public TestPart As clsParticulas


Public Function FtoDW(f As Single) As Long
    Dim buf As D3DXBuffer
    Dim l As Long
    Set buf = D3DX.CreateBuffer(4)
    D3DX.BufferSetData buf, 0, 4, 1, f
    D3DX.BufferGetData buf, 0, 4, 1, l
    FtoDW = l
End Function
