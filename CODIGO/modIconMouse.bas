Attribute VB_Name = "modIconMouse"
Option Explicit
Public Enum ModosDeStretch
    BlackOnWhite = 1
    WhiteOnBlack = 2
    ColorOnColor = 3
    Halftone = 4
    Desconocida = 5
End Enum

Private Type PICTDESC
    cbSizeOfStruct As Long
    picType As Long
    hgdiObj As Long
    hPalOrXYExt As Long
End Type

Private Type IID
    data1 As Long
    data2 As Integer
    Data3 As Integer
    Data4(0 To 7)  As Byte
End Type

Private Declare Sub OleCreatePictureIndirect Lib "oleaut32.dll" (lpPictDesc As PICTDESC, riid As IID, ByVal fOwn As Boolean, lplpvObj As Object)
    
Private Type pvICONINFO
    fIcon As Long
    xHotspot As Long
    yHotspot As Long
    hbmMask As Long
    hbmColor As Long
End Type

Private Type pvRECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function CreateIconIndirect Lib "user32" (piconinfo As pvICONINFO) As Long

Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long

Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lppvRECT As pvRECT, ByVal hBrush As Long) As Long

Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal color As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal color As Long) As Long

Const DI_MASK = &H1
Const DI_IMAGE = &H2
Public Sub pvBMPaICO(srchDC As Long, ByVal X As Integer, ByVal Y As Integer, ByVal hImagen As Long, ByVal hMask As Long, ByVal hdc As Long, ByVal ScrDC As Long, ByVal MaskColor As Long, ByVal ModoDeStretch As Byte)
Dim r As pvRECT, hBr As Long
Dim hOldPal As Long, hDC_Copia As Long
Dim TmpBMP As Long

    r.Bottom = 32
    r.Right = 32
    
    SetStretchBltMode hdc, ModoDeStretch
            
    ' Dibujo la mascara
    
    If MaskColor = -1 Then    ' No hay transparencia
        
        ' Selecciono la mascara...
        SelectObject hdc, hMask
        
        ' ... y la lleno con negro (opaco)
        hBr = CreateSolidBrush(&H0)
        FillRect hdc, r, hBr
        DeleteObject hBr
    
        ' Selecciono la imagen
        SelectObject hdc, hImagen
    
        StretchBlt hdc, 0, 0, 32, 32, srchDC, X, Y, 32, 32, vbSrcCopy
 
               
    Else

        ' Creo un DC y un bitmap para
        ' copiar la imagen. Esto lo
        ' debo hacer porque si el bitmap
        ' es DIB no pasa a B&N usando
        ' los colores de fondo y texto.
        
        hDC_Copia = CreateCompatibleDC(ScrDC)
        
        SetStretchBltMode hDC_Copia, ModoDeStretch
        
        TmpBMP = CreateCompatibleBitmap(ScrDC, 32, 32)
                
        ' Hago la copia del bitmap
        SelectObject hDC_Copia, TmpBMP
        
        StretchBlt hDC_Copia, 0, 0, 32, 32, srchDC, X, Y, 32, 32, vbSrcCopy
            
              
        ' De ahora en mas utilizo la copia
        ' de la que ya a sido modificado su
        ' tama~o
              
        ' ---- Creo la mascara -----
        
        ' Selecciono la mascara en el DC
        SelectObject hdc, hMask
        
        ' Seteo el color de fondo con
        ' el color de mascara.
        SetBkColor hDC_Copia, MaskColor
        SetTextColor hDC_Copia, vbWhite
        
        ' Al copiar windows transforma en blanco
        ' todos los pixel con el color de fondo
        ' y en negro el resto
        BitBlt hdc, 0, 0, 32, 32, hDC_Copia, 0, 0, vbSrcCopy
          
        SelectObject hdc, hImagen
        SelectObject hDC_Copia, hMask

        hBr = CreateSolidBrush(&H0)
        FillRect hdc, r, hBr
        DeleteObject hBr
        
        ' Copio la mascara y luego la imagen
        BitBlt hdc, 0, 0, 32, 32, hDC_Copia, 0, 0, vbNotSrcCopy
        BitBlt hdc, 0, 0, 32, 32, srchDC, 0, 0, vbSrcAnd
            
            
        DeleteDC hDC_Copia
        DeleteObject TmpBMP
        
    End If

End Sub

Public Function HandleToPicture(ByVal hGDIHandle As Long, ByVal ObjectType As PictureTypeConstants, Optional ByVal hPal As Long = 0) As StdPicture
Dim ipic As IPicture, picdes As PICTDESC, iidIPicture As IID
    
    ' Fill picture description
    picdes.cbSizeOfStruct = Len(picdes)
    picdes.picType = ObjectType
    picdes.hgdiObj = hGDIHandle
    picdes.hPalOrXYExt = hPal
    
    ' IPictureDisp {7BF80981-BF32-101A-8BBB-00AA00300CAB}
    iidIPicture.data1 = &H7BF80981

    iidIPicture.data2 = &HBF32

    iidIPicture.Data3 = &H101A

    iidIPicture.Data4(0) = &H8B

    iidIPicture.Data4(1) = &HBB

    iidIPicture.Data4(2) = &H0

    iidIPicture.Data4(3) = &HAA

    iidIPicture.Data4(4) = &H0

    iidIPicture.Data4(5) = &H30

    iidIPicture.Data4(6) = &HC

    iidIPicture.Data4(7) = &HAB

    
    ' Crea el objeto con el handle
    OleCreatePictureIndirect picdes, iidIPicture, True, ipic
    
    Set HandleToPicture = ipic
        
End Function
Public Function GetIcon(ByVal srchDC As Long, ByVal srcX As Integer, ByVal srcY As Integer, Optional ModoDeStretch As ModosDeStretch = Halftone, Optional CrearCursor As Boolean = False, Optional MaskColor As Long = -1) As StdPicture
Dim hIcon As Long, IconPict As StdPicture
Dim ScreenDC As Long, BitmapDC As Long
Dim hMask As Long, hImagen As Long
Dim hIcn As Long, II As pvICONINFO

    On Error Resume Next

   
    ScreenDC = GetWindowDC(0&)
    BitmapDC = CreateCompatibleDC(ScreenDC)
    
    hImagen = CreateCompatibleBitmap(ScreenDC, 32, 32)
    hMask = CreateBitmap(32, 32, 1, 1, ByVal 0&)
    
      
        pvBMPaICO srchDC, srcX, srcY, hImagen, hMask, BitmapDC, ScreenDC, MaskColor, ModoDeStretch

    
    DeleteDC BitmapDC
    ReleaseDC 0&, ScreenDC
    
    II.fIcon = CrearCursor
    II.hbmColor = hImagen
    II.hbmMask = hMask
    
    hIcon = CreateIconIndirect(II)

    
    Set IconPict = HandleToPicture(hIcon, vbPicTypeIcon)
     
    If IconPict Is Nothing Then
        
        DeleteObject hIcn
        Set GetIcon = Nothing
        
    Else
        
        Set GetIcon = IconPict
        
    End If
    
     
End Function


