VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSurfaceManStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
' clsSurfaceManStatic.cls - Inherits from clsSurfaceManager. Is designed to load
' surfaces at startup, and never deallocating them.
' This grants high performance can use a lot of RAM.
'
' Developed by Maraxus (Juan Martín Sotuyo Dodero - juansotuyo@hotmail.com)
' Last Modify Date: 3/06/2006
'**************************************************************

'**************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'**************************************************************

Option Explicit

'Inherit from the surface manager
Implements clsSurfaceManager

'Number of buckets in our hash table. Must be a nice prime number.
Const HASH_TABLE_SIZE As Long = 541

Private Type SURFACE_ENTRY_STATIC
    fileIndex As Long
    Surface As Direct3DTexture8
End Type

Private Type HashBucket
    surfaceCount As Integer
    SurfaceEntry() As SURFACE_ENTRY_STATIC
End Type

Private surfaceList(HASH_TABLE_SIZE - 1) As HashBucket

Private DeviceD3D As Direct3DDevice8
Private DirectD3X As D3DX8

Private useVideoMemory As Boolean

Private Const DIB_PAL_COLORS As Long = 1
Private Const DIB_RGB_COLORS As Long = 0
    
Private Sub Class_Initialize()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/06/2006
'
'**************************************************************
End Sub

Private Sub Class_Terminate()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/06/2006
'Clean up
'**************************************************************
    Dim i As Long
    Dim j As Long
    
    'Destroy every surface in memory
    For i = 0 To HASH_TABLE_SIZE - 1
        With surfaceList(i)
            For j = 1 To .surfaceCount
                Set .SurfaceEntry(j).Surface = Nothing
                .SurfaceEntry(j).fileIndex = 0
            Next j
            surfaceList(i).surfaceCount = 0
            'Destroy the arrays
            Erase .SurfaceEntry
        End With
    Next i
End Sub
Private Sub clsSurfaceManager_Limpiar()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/06/2006
'Clean up
'**************************************************************
    Dim i As Long
    Dim j As Long
    
    'Destroy every surface in memory
    For i = 0 To HASH_TABLE_SIZE - 1
        With surfaceList(i)
            For j = 1 To .surfaceCount
                Set .SurfaceEntry(j).Surface = Nothing
                .SurfaceEntry(j).fileIndex = 0
            Next j
            surfaceList(i).surfaceCount = 0
            'Destroy the arrays
            Erase .SurfaceEntry
        End With
    Next i
End Sub
Private Sub clsSurfaceManager_Initialize(ByRef D3DD As Direct3DDevice8, ByRef D3DX As D3DX8, ByVal videoMemory As Boolean, ByVal graphicPath As String, Optional ByVal maxMemoryUsageInMb As Long = -1&)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/06/2006
'
'**************************************************************
    Set DeviceD3D = D3DD
    Set DirectD3X = D3DX
    
    useVideoMemory = videoMemory
    
    Call LoadSurfaces(graphicPath)
End Sub

Private Property Get clsSurfaceManager_Surface(ByVal fileIndex As Long) As Direct3DTexture8
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/06/2006
'Retrieves the requested texture
'**************************************************************
On Error GoTo ErrHandler:
    Dim i As Long
    
    ' Search for it and return it
    With surfaceList(fileIndex Mod HASH_TABLE_SIZE)
        For i = 1 To .surfaceCount
            If .SurfaceEntry(i).fileIndex = fileIndex Then
                Set clsSurfaceManager_Surface = .SurfaceEntry(i).Surface
                Exit Property
            End If
        Next i
    End With
    
    'We have a missing bmp!
    MsgBox "No se encuentra el archivo " & CStr(fileIndex) & ".bmp. Reinstale el juego, " _
        & "y si el problema persiste contactese con los adminsitradores", vbOKOnly Or vbCritical Or vbExclamation, "Error"
    
    End
Exit Property

ErrHandler:
    MsgBox "Un error inesperado ocurrió a intentar cargar el archivo " & CStr(fileIndex) & ".bmp. & vbcrlf" _
        & "Error: " & CStr(Err.Number), vbOKOnly Or vbCritical Or vbExclamation, "Error"
    
    End
End Property

Private Sub LoadSurfaces(ByVal ResourcePath As String)
'**************************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modify Date: 12/02/2007
'Loads all surfaces in random order and then sorts them
'**************************************************************
On Error Resume Next
    Dim surfaceIndex As Long
    Dim Data() As Byte
    Dim i As Long
    
    While GetNext_File(ResourcePath, i, Data(), surfaceIndex)
        Call LoadSurface(Data(), surfaceIndex)
    Wend
End Sub

Private Sub LoadSurface(ByRef Data() As Byte, ByVal fileIndex As Long)
'**************************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modify Date: 11/30/2007
'Loads the surface named fileIndex + ".bmp" and inserts it to the
'surface list in the listIndex position
'**************************************************************
On Error GoTo ErrHandler

    Dim newSurface As SURFACE_ENTRY_STATIC
    Dim Tamaño As Long
    Dim pDesc As D3DSURFACE_DESC
    'ddsd.lWidth = bmpInfo.bmiHeader.biWidth
    'ddsd.lHeight = bmpInfo.bmiHeader.biHeight
    Tamaño = UBound(Data) + 1
    
    With newSurface
        .fileIndex = fileIndex
        
        'Load surface
        Set .Surface = DirectD3X.CreateTextureFromFileInMemoryEx(DeviceD3D, Data(0), Tamaño, 0, 0, D3DX_DEFAULT, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, ColorTransparenteDX, ByVal 0, ByVal 0)
        .Surface.GetLevelDesc 0, pDesc
        SurfaceSize(fileIndex).X = pDesc.Width
        SurfaceSize(fileIndex).Y = pDesc.Height
    End With
    
    'Store the surface in the given index (it MUST be empty or data will be lost)
    With surfaceList(fileIndex Mod HASH_TABLE_SIZE)
        .surfaceCount = .surfaceCount + 1
        ReDim Preserve .SurfaceEntry(1 To .surfaceCount) As SURFACE_ENTRY_STATIC
        
        .SurfaceEntry(.surfaceCount) = newSurface
    End With
Exit Sub

ErrHandler:
    If Err.Number = 1 Then
        ' No video memory? Try system memory! We want all surfaces loaded!
        Resume
    End If
    
    MsgBox "Un error inesperado ocurrió al intentar cargar el gráfico " & CStr(fileIndex) & ".BMP" & ". " & vbCrLf & _
            "El código de error es " & CStr(Err.Number) & " - " & Err.Description & vbCrLf & vbCrLf & "Copia este mensaje y notifica a los administradores.", _
            vbOKOnly Or vbCritical Or vbExclamation, "Error"
        End
End Sub
