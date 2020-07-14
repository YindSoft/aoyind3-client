VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBufferMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************************
'CBufferMan - ImperiumAO - v1.4.5 R5
'
'Sound buffer manager based on Maraxus's texture manager.
'
'*****************************************************************
'Respective portions copyrighted by contributors listed below.
'
'This library is free software; you can redistribute it and/or
'modify it under the terms of the GNU Lesser General Public
'License as published by the Free Software Foundation version 2.1 of
'the License
'
'This library is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
'Lesser General Public License for more details.
'
'You should have received a copy of the GNU Lesser General Public
'License along with this library; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA

'*****************************************************************
'Maraxus (Juan Martín Sotuyo Dodero - juansotuyo@hotmail.com)
'   - First Relase (as CTextureMan)
'Barrin (Augusto José Rando - barrin@imperiumao.com.ar)
'   - Modified to use sound buffers
'   - Modified to copy buffers when already streaming (multi-buffering)
'*****************************************************************

Option Explicit


Private Type WAV_DB_ENTRY
    FileName As Integer
    UltimoAcceso As Long
    buffer As DirectSoundSecondaryBuffer8
End Type

Private Type WAVETYPE
    strHead As String * 12
    strFormatID As String * 4
    lngChunkSize As Long
    intFormat As Integer
    intChannels As Integer
    lngSamplesPerSec As Long
    lngAvgBytesPerSec As Long
    intBlockAlign As Integer
    intBitsPerSample As Integer
End Type

Private mBuffers() As WAV_DB_ENTRY
Private mDS As DirectSound8

Private mMaxEntries As Integer
Private mCantidadBuffers As Integer

Private WavPath As String

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Private Sub Class_Initialize()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Modified by Augusto José Rando
'Last Modify Date: 5/04/2005
'
'**************************************************************

End Sub

Private Sub Class_Terminate()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/04/2005
'
'**************************************************************
On Error Resume Next
    Dim i As Long
    Dim j As Long
    
    For i = 1 To mCantidadBuffers
        Set mBuffers(i).buffer = Nothing
    Next i
    
    Erase mBuffers
    
End Sub

Public Property Let Path(ByVal inString As String)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/04/2005
'
'**************************************************************
    WavPath = inString
End Property

Public Sub GetBuffer(ByVal FileName As Integer, ByRef tBuff As DirectSoundSecondaryBuffer8)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Modified by Augusto José Rando
'Last Modify Date: 9/05/2005
'
'**************************************************************

    Dim IndiceObtenido As Integer
    'Dim tCap As DSBCAPS
    IndiceObtenido = ObtenerIndice(FileName)
    
    If IndiceObtenido > 0 Then
            
        If Not mBuffers(IndiceObtenido).buffer Is Nothing Then
            'If mBuffers(IndiceObtenido).buffer.GetStatus = DSBSTATUS_PLAYING Then
                'Call mBuffers(IndiceObtenido).buffer.GetCaps(tCap)
                'Call CopyMemory(ObjPtr(tBuff), ByVal ObjPtr(mBuffers(IndiceObtenido).buffer), tCap.lBufferBytes)
                'Set tBuff = mDS.DuplicateSoundBuffer(ByVal mBuffers(IndiceObtenido).buffer)
                
                'Call tBuff.GetCaps(tCap)
                'Debug.Print tCap.lBufferBytes
                
            'Else
                'Devuelvo un buffer con el sonido cargado
                Set tBuff = mBuffers(IndiceObtenido).buffer
            'End If
        Else
            If CreateBufferFromFile_Ex(FileName, IndiceObtenido) Then
                Set tBuff = mBuffers(IndiceObtenido).buffer
            End If
        End If
                
        'Ultimo acceso
        mBuffers(IndiceObtenido).UltimoAcceso = GetTickCount
            
    Else    'Sonido no cargado
        GoTo CrearNuevoBuffer
    End If
    
Exit Sub
    
CrearNuevoBuffer:
    'Vemos si puedo agregar uno a la lista
    If mMaxEntries = mCantidadBuffers Then
        'Sacamos el que hace más que no usamos, y utilizamos el slot
        IndiceObtenido = CrearBuffer(FileName, BorraMenosUsado())
        Set tBuff = mBuffers(IndiceObtenido).buffer
    Else
        'Agrego un buffer nueva a la lista
        IndiceObtenido = CrearBuffer(FileName, IndiceObtenido)
        Set tBuff = mBuffers(IndiceObtenido).buffer
    End If
    
End Sub

Private Function ObtenerIndice(ByVal FileName As Integer) As Integer
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/04/2005
'Busqueda binaria para hallar el buffer deseado
'**************************************************************
    Dim max As Integer  'Max index
    Dim min As Integer  'Min index
    Dim mid As Integer  'Middle index
    
    min = 1
    mid = 1
    max = mCantidadBuffers
    
    Do While min <= max
        mid = (min + max) / 2
        If FileName < mBuffers(mid).FileName Then
            'El índice no existe
            max = mid - 1
        ElseIf FileName > mBuffers(mid).FileName Then
            'El índice no existe
            min = mid + 1
        Else
            ObtenerIndice = mid
            Exit Function
        End If
    Loop
    
    'Maraxus - usado para binary insertion
    ObtenerIndice = Not mid
    
End Function

Public Function Init(ByRef DS8 As DirectSound8, ByVal MaxEntries As Integer) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/04/2005
'
'**************************************************************
    
    mMaxEntries = MaxEntries
    
    If mMaxEntries < 1 Then 'por lo menos 1 sonido
        Exit Function
    End If
    
    mCantidadBuffers = 0
    
    'Seteamos el objeto
    Set mDS = DS8
    
    Init = True
End Function

Public Sub BorrarTodo()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/04/2005
'Vacia la lista de sonidos
'**************************************************************
    Dim i As Long, j As Long
    
    For i = 1 To mCantidadBuffers
        Set mBuffers(i).buffer = Nothing
    Next i
    
    ReDim mBuffers(0)
    mCantidadBuffers = 0
End Sub

Public Sub DetenerTodo()
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: 5/04/2005
'Detiene todos los sonidos
'**************************************************************
    Dim i As Long
    Dim j As Long
    
    For i = 1 To mCantidadBuffers
        If Not mBuffers(i).buffer Is Nothing Then
            If (mBuffers(i).buffer.GetStatus = DSBSTATUS_PLAYING + DSBSTATUS_LOOPING) Or (mBuffers(i).buffer.GetStatus = DSBSTATUS_PLAYING) Then
                mBuffers(i).buffer.Stop
                mBuffers(i).buffer.SetCurrentPosition 0
            End If
        End If
                
    Next i
    
End Sub

Public Sub ModificarVolumenTodo(ByVal Volume As Long)
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: 5/04/2005
'Detiene todos los sonidos
'**************************************************************
    Dim i As Long
    Dim j As Long
    
    For i = 1 To mCantidadBuffers
        If Not mBuffers(i).buffer Is Nothing Then
            If (mBuffers(i).buffer.GetStatus = DSBSTATUS_LOOPING + DSBSTATUS_PLAYING) Or (mBuffers(i).buffer.GetStatus = DSBSTATUS_PLAYING) Then
                mBuffers(i).buffer.SetVolume Volume
            End If
        End If
    Next i
    
End Sub

Public Function Borrar(ByVal FileName As Integer) As Integer
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/04/2005
'Borra un sonido
'**************************************************************
    Dim indice As Long
    
    'Obtenemos el ìndice
    indice = ObtenerIndice(FileName)
    If indice < 0 Then Exit Function
    
    'Lo eliminamos
    With mBuffers(indice)
        Set .buffer = Nothing
        .FileName = 0
        .UltimoAcceso = 0
    End With
    
    'Actualizamos el número de sonidos
    mCantidadBuffers = mCantidadBuffers - 1
    
    'Movemos para atrás el resto de la lista
    For indice = indice To mCantidadBuffers
        mBuffers(indice) = mBuffers(indice + 1)
    Next indice
    
    'Redimencionamos la lista
    ReDim Preserve mBuffers(1 To mCantidadBuffers)
End Function

Private Function CrearBuffer(ByVal Archivo As Integer, ByVal Index As Integer) As Integer
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/04/2005
'
'**************************************************************
On Error GoTo errhandler
        
    If Index < 0 Then
        Index = Not Index   ' Obtenemos el índice donde debe ser insertado
        ReDim Preserve mBuffers(1 To mCantidadBuffers + 1) As WAV_DB_ENTRY    ' Incrementamos la lista
        If Index < mCantidadBuffers + 1 Then
            Dim i As Long
            For i = mCantidadBuffers To Index Step -1
                mBuffers(i + 1) = mBuffers(i)
            Next i
        End If
    End If
    
    With mBuffers(Index)
        'Nombre
        .FileName = Archivo
        
        'Ultimo acceso
        .UltimoAcceso = GetTickCount
        
        Call CreateBufferFromFile_Ex(Archivo, Index)
        
    End With
    
    'Aumentamos la cantidad de sonidos
    mCantidadBuffers = mCantidadBuffers + 1
        
    'Devolvemos el ìndice en que lo cargamos
    CrearBuffer = Index
Exit Function

errhandler:
    
End Function

Private Function CreateBufferFromFile(ByVal file_name As Integer, ByVal buffer_index As Integer) As Boolean
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: 3/04/2007
'Creación de buffers desde archivo (despreciado)
'**************************************************************

Dim dsbd As DSBUFFERDESC, dsbcap As DSBCAPS

dsbd.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME
dsbd.fxFormat.nFormatTag = WAVE_FORMAT_PCM
dsbd.fxFormat.nChannels = 2
dsbd.fxFormat.lSamplesPerSec = 22050
dsbd.fxFormat.nBitsPerSample = 16
dsbd.fxFormat.nBlockAlign = dsbd.fxFormat.nBitsPerSample / 8 * dsbd.fxFormat.nChannels
dsbd.fxFormat.lAvgBytesPerSec = dsbd.fxFormat.lSamplesPerSec * dsbd.fxFormat.nBlockAlign

If Extract_File(Wav, WavPath, file_name & ".wav", Windows_Temp_Dir, False) Then
    Set mBuffers(buffer_index).buffer = mDS.CreateSoundBufferFromFile(Windows_Temp_Dir & file_name & ".wav", dsbd)
    Delete_File Windows_Temp_Dir & file_name & ".wav"
    CreateBufferFromFile = True
End If

Exit Function

errhandler:
    'Ocurrió un error, la causa más probable es que el archivo no exista
    If General_File_Exists(Windows_Temp_Dir & file_name & ".wav", vbNormal) Then
        Delete_File Windows_Temp_Dir & file_name & ".wav"
    End If

End Function

Private Function CreateBufferFromFile_Ex(ByVal file_name As Integer, ByVal buffer_index) As Boolean
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: 3/04/2007
'Creación de buffers desde memoria
'**************************************************************

On Error GoTo ErrorHandler

Dim btArr() As Byte, gudtHeader As WAVETYPE

Dim i As Long, lngChunkSize As Long

Dim tB(1 To 4) As Byte

Dim glngChunkSize As Long
Dim lngOffset As Long

Dim udtBufferDesc As DSBUFFERDESC

If Get_File_Data(DirRecursos & "Wav.AO", file_name & ".wav", btArr) Then
    
    lngOffset = LenB(gudtHeader)
    
    Call CopyMemory(gudtHeader, btArr(0), lngOffset)
    
    'Only PCM
    If gudtHeader.intFormat <> WAVE_FORMAT_PCM Then Exit Function
    
    lngOffset = 36
    
    For i = lngOffset To UBound(btArr)
        tB(1) = btArr(i)
        tB(2) = btArr(i + 1)
        tB(3) = btArr(i + 2)
        tB(4) = btArr(i + 3)
        
        If StrConv(tB(), vbUnicode) = "data" Then
            Exit For
        End If
        
    Next i
    
    lngOffset = i + 4
    
    Call CopyMemory(ByVal VarPtr(glngChunkSize), btArr(lngOffset), ByVal 4)
    
    lngOffset = lngOffset + 4
    
    'Set the Wave Format
    With udtBufferDesc.fxFormat
        .nFormatTag = gudtHeader.intFormat
        .nChannels = gudtHeader.intChannels
        .lSamplesPerSec = gudtHeader.lngSamplesPerSec
        .nBitsPerSample = gudtHeader.intBitsPerSample
        .nBlockAlign = gudtHeader.intBlockAlign '/ 8 * gudtHeader.intChannels
        .lAvgBytesPerSec = gudtHeader.lngAvgBytesPerSec '* .nBlockAlign
        .nsize = gudtHeader.lngChunkSize
    End With
            
    'Create the buffer
    udtBufferDesc.lBufferBytes = glngChunkSize
    udtBufferDesc.lFlags = DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME
    
    Set mBuffers(buffer_index).buffer = mDS.CreateSoundBuffer(udtBufferDesc)
                
    'Load the buffer with data
    mBuffers(buffer_index).buffer.WriteBuffer 0, glngChunkSize, btArr(lngOffset), DSBLOCK_ENTIREBUFFER
        
    CreateBufferFromFile_Ex = True
End If

Exit Function

ErrorHandler:
    'Ocurrió un error, la causa más probable es que el archivo no exista
    'If General_File_Exists(Windows_Temp_Dir & file_name & ".wav", vbNormal) Then
    '    Delete_File Windows_Temp_Dir & file_name & ".wav"
    'End If

End Function

Private Function BorraMenosUsado() As Integer
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/04/2005
'Borra el sonido menos usado. Devuelve el index del borrado para ser reutilizado
'NO redimenciona la lista, estamos forzando a sacar uno para meter otro en su lugar
'**************************************************************
    Dim Valor As Long
    Dim i As Long
    
    'Inicializamos todo
    Valor = mBuffers(1).UltimoAcceso
    BorraMenosUsado = 1
    
    'Buscamos cual es el que lleva más tiempo sin ser utilizado
    For i = 1 To mCantidadBuffers
        If mBuffers(i).UltimoAcceso < Valor And (Not PlayingCheck(0, i)) Then
            Valor = mBuffers(i).UltimoAcceso
            BorraMenosUsado = i
        End If
    Next i
    
    'Disminuimos el contador
    mCantidadBuffers = mCantidadBuffers - 1
    
    'Borramos los buffers
    Set mBuffers(BorraMenosUsado).buffer = Nothing
    
    mBuffers(BorraMenosUsado).FileName = 0
    mBuffers(BorraMenosUsado).UltimoAcceso = 0
    
    ' Redimensionamos el array
    ReDim Preserve mBuffers(mCantidadBuffers)
    
End Function

Public Property Get MaxEntries() As Integer
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/04/2005
'
'**************************************************************
    MaxEntries = mMaxEntries
End Property

Public Property Let MaxEntries(ByVal vNewValue As Integer)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/04/2005
'
'**************************************************************
    mMaxEntries = vNewValue
End Property

Public Property Get CantidadBuffers() As Integer
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/04/2005
'
'**************************************************************
    CantidadBuffers = mCantidadBuffers
End Property

Public Sub BorraTimerProc()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/04/2005
'Cada un minuto revisamos que buffers se pueden eliminar
'**************************************************************
    On Error Resume Next
    
    Dim loopc As Long
    Dim offset As Integer
    Dim TiempoActual As Long
    
    TiempoActual = GetTickCount
    
    For loopc = 1 To mCantidadBuffers
        'Revisar cuales usamos
        If mBuffers(loopc).UltimoAcceso > TiempoActual - 60000 Then
            'Si el offset está seteado, movemos para atrás la lista
            If offset Then
                mBuffers(loopc - offset) = mBuffers(loopc)
            End If
        Else
            'Eliminar de memoria
            Set mBuffers(loopc).buffer = Nothing
            
            'Clear variables
            mBuffers(loopc).FileName = 0
            mBuffers(loopc).UltimoAcceso = 0
            
            offset = offset + 1
        End If
    Next loopc
    
    'Actualizar el número de sonidos
    mCantidadBuffers = mCantidadBuffers - offset
End Sub

Public Function CargarBuffer(ByVal FileName As Integer) As Boolean
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: 17/08/2005
'
'**************************************************************

    Dim Index As Integer
    Index = ObtenerIndice(FileName)
    
    If Index > 0 Then
        CargarBuffer = True
    Else    'Sonido no cargado
        'Vemos si puedo agregar uno a la lista
        If mMaxEntries = mCantidadBuffers Then
            'Sacamos el que hace más que no usamos, y utilizamos el slot
            Call CrearBuffer(FileName, BorraMenosUsado())
            CargarBuffer = True
        Else
            'Agrego un buffer nueva a la lista
            Call CrearBuffer(FileName, Index)
            CargarBuffer = True
        End If
    End If
    
End Function

Public Function Reproducir(ByVal file_name As Integer, play_format As CONST_DSBPLAYFLAGS, ByVal Volume As Long, ByVal Pan As Long) As Boolean
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: 17/08/2005
'
'**************************************************************

Static Sound_Buffer As DirectSoundSecondaryBuffer8

On Error GoTo ErrorHandler

    Call GetBuffer(file_name, Sound_Buffer)
    If Sound_Buffer Is Nothing Then Exit Function
    
    'Esto solo debería pasar si ya hay MAX_BUFFERS en reproducción
    If Sound_Buffer.GetStatus = DSBSTATUS_PLAYING Then Exit Function
    
    Sound_Buffer.SetPan Pan
    Sound_Buffer.SetVolume Volume
            
    Sound_Buffer.Play play_format
    
    Reproducir = True

    Exit Function

ErrorHandler:
    MsgBox Err.Description
    Reproducir = False

End Function

Public Function Detener(ByVal file_name As Integer) As Boolean
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: 17/08/2005
'
'**************************************************************

On Error GoTo ErrorHandler

Dim IndiceObtenido As Integer, i As Long

    IndiceObtenido = ObtenerIndice(file_name)
    If IndiceObtenido < 0 Then Exit Function
    
    If Not mBuffers(IndiceObtenido).buffer Is Nothing Then
        mBuffers(IndiceObtenido).buffer.Stop
        mBuffers(IndiceObtenido).buffer.SetCurrentPosition 0
    End If
        
    Detener = True

    Exit Function

ErrorHandler:
    Detener = False

End Function

Public Function CambiarVolumen(ByVal file_name As Integer, ByVal volumen As Long) As Boolean
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: 17/08/2005
'Returns true if successfully changed volume of a playing buffer
'**************************************************************

On Error GoTo ErrorHandler

Dim IndiceObtenido As Integer, i As Long, bFlag As Boolean

    IndiceObtenido = ObtenerIndice(file_name)
    If IndiceObtenido < 0 Then Exit Function
    
    If Not mBuffers(IndiceObtenido).buffer Is Nothing Then
        mBuffers(IndiceObtenido).buffer.SetVolume volumen
        If (mBuffers(IndiceObtenido).buffer.GetStatus = DSBSTATUS_PLAYING Or mBuffers(IndiceObtenido).buffer.GetStatus = DSBSTATUS_PLAYING + DSBSTATUS_LOOPING) Then bFlag = True
    End If
        
    CambiarVolumen = bFlag

    Exit Function

ErrorHandler:
    CambiarVolumen = False

End Function

Public Function CambiarPan(ByVal file_name As Integer, ByVal Pan As Long) As Boolean
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: 17/08/2005
'Returns true if successfully changed pan of a playing buffer
'**************************************************************

On Error GoTo ErrorHandler

Dim IndiceObtenido As Integer, i As Long, bFlag As Boolean

    IndiceObtenido = ObtenerIndice(file_name)
    If IndiceObtenido < 0 Then Exit Function
    
    If Not mBuffers(IndiceObtenido).buffer Is Nothing Then
        mBuffers(IndiceObtenido).buffer.SetPan Pan
        If (mBuffers(IndiceObtenido).buffer.GetStatus = DSBSTATUS_PLAYING Or mBuffers(IndiceObtenido).buffer.GetStatus = DSBSTATUS_PLAYING + DSBSTATUS_LOOPING) Then bFlag = True
    End If
        
    CambiarPan = bFlag

    Exit Function

ErrorHandler:
    CambiarPan = False

End Function

Public Function PlayingCheck(ByVal file_name As Integer, Optional ByVal IndiceObtenido As Integer = -1) As Boolean
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: 17/08/2005
'Returns true if playing
'**************************************************************

On Error GoTo ErrorHandler

Dim i As Long, bFlag As Boolean

    If IndiceObtenido = -1 Then
        IndiceObtenido = ObtenerIndice(file_name)
        If IndiceObtenido < 0 Then Exit Function
    End If
    
    If Not mBuffers(IndiceObtenido).buffer Is Nothing Then
        If (mBuffers(IndiceObtenido).buffer.GetStatus = DSBSTATUS_PLAYING + DSBSTATUS_LOOPING) Or (mBuffers(IndiceObtenido).buffer.GetStatus = DSBSTATUS_PLAYING) Then
            PlayingCheck = True
            Exit Function
        End If
    End If

    Exit Function

ErrorHandler:
    PlayingCheck = False

End Function

