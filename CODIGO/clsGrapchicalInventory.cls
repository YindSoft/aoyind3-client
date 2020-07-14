VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGrapchicalInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************
'Graphical inventory for ORE 0.4.1
'Designed and implemented by Juan Martín Sotuyo Dodero (Maraxus)
'juansotuyo@hotmail.com
'*****************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'**************************************************************************

Option Explicit

Private InventoryOffset As Long             'Number of lines we scrolled down from topmost
Private InvSelectedItem As Long             'Currently selected item

Private ownTilePixelWidth As Integer
Private ownTilePixelHeight As Integer
Private startdX As Integer
Private startdY As Integer

Private ShowValue As Boolean
Private ImgContainer As Boolean
Private EsMain As Boolean

Private Type pvRECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Inventory() As Inventory    'User's inventory

Private WithEvents InventoryWindow As PictureBox    'Placeholder where to render the inventory
Attribute InventoryWindow.VB_VarHelpID = -1

Private Declare Function SetDIBitsToDevice Lib "gdi32.dll" _
    (ByVal hdc As Long, ByVal XDest As Long, ByVal YDest As Long, _
    ByVal dwWidth As Long, ByVal dwHeight As Long, ByVal xSrc As Long, _
    ByVal ySrc As Long, ByVal uStartScan As Long, ByVal CScanLine As Long, _
    lpBits As Any, lpbmi As BITMAPINFO, ByVal fuColorUse As Long) As Long
    
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lppvRECT As pvRECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long


Private Const DIB_PAL_COLORS As Long = 1
Private Const DIB_RGB_COLORS As Long = 0

Private Graficos As Dictionary
Private FONDO As Long

Public OFFSETX As Integer
Public OFFSETY As Integer

Public DropX As Integer
Public DropY As Integer

Private Sub Class_Initialize()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'
'***************************************************
Set Graficos = New Dictionary
End Sub

Private Sub Class_Terminate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'
'***************************************************
Dim i As Integer
For i = 1 To Graficos.Count
    DeleteDC (Graficos(i))
Next i
If FONDO <> 0 Then
    DeleteDC (FONDO)
End If
End Sub

Public Sub ScrollInventory(ByVal Up As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 14/01/2010
'Scrolls the graphical inventory up / down
'14/01/2010: ZaMa - Now redraws inventory only if needed
'***************************************************

    'Only move if there is something to show up /down
    If CanScroll(Up) Then
        InventoryOffset = InventoryOffset + IIf(Up, 1, -1)
        
        'Redraw the inventory if needed
        DrawInventory
    End If

End Sub

Public Function CanScroll(ByVal Up As Boolean)
'***************************************************
'Author: ZaMa
'Last Modify Date: 14/01/2010
'Determines wether inventory can be scrolled up / down
'***************************************************
    If Up Then
        CanScroll = (InventoryOffset + (InventoryWindow.ScaleHeight \ ownTilePixelHeight)) * _
                    (InventoryWindow.ScaleWidth \ ownTilePixelWidth) < UBound(Inventory)
    Else
        CanScroll = InventoryOffset > 0
    End If
    
End Function

Public Sub SelectGold()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Sets the gold as the currently selected item
'***************************************************
    Dim prevSelectedItem As Long
    
    'Store preivously selected item
    prevSelectedItem = InvSelectedItem
    
    'Select the gold
    InvSelectedItem = FLAGORO
    
    'Redraw old item to deselect it
    If prevSelectedItem <> FLAGORO Then
        Call DrawInventorySlot(prevSelectedItem)
    End If
End Sub

Public Sub DeselectItem()
'***************************************************
'Author: ZaMa
'Last Modify Date: 30/11/2009
'Deselect the currently selected item
'***************************************************
    If InvSelectedItem = 0 Then Exit Sub
    
    Dim ClearSlot As Byte
    
    ClearSlot = InvSelectedItem
    
    'Select nothing
    InvSelectedItem = 0
    
    'Redraw old item to deselect it
    Call DrawInventorySlot(ClearSlot)

End Sub

Public Sub ChangeSlotItemAmount(ByVal slot As Byte, ByVal Amount As Long)
'***************************************************
'Author: ZaMa
'Created Date: 30/11/2009
'Last Modify Date: 09/12/2009 (Budi)
'Change the amount of the required item, and update the slot.
'***************************************************
    
    Inventory(slot).Amount = Amount
    
    'Update Amount
    Call DrawInventorySlot(slot)

End Sub

Public Property Get SelectedItem() As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item index
'***************************************************
    SelectedItem = InvSelectedItem
End Property

Public Property Get MaxHit(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the max hit of the selected item
'***************************************************
    MaxHit = Inventory(slot).MaxHit
End Property

Public Property Get MinHit(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the min hit of the selected item
'***************************************************
    MinHit = Inventory(slot).MinHit
End Property

Public Property Get MaxDef(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the defense of the selected item
'***************************************************
    MaxDef = Inventory(slot).MaxDef
End Property

Public Property Get MinDef(ByVal slot As Byte) As Integer
'***************************************************
'Author: Budi
'Last Modify Date: 02/1209
'Retrieves the defense of the selected item
'***************************************************
    MinDef = Inventory(slot).MinDef
End Property

Public Property Get GrhIndex(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the grh index of the selected item
'***************************************************
    GrhIndex = Inventory(slot).GrhIndex
End Property

Public Property Get Valor(ByVal slot As Byte) As Single
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the value of the selected item
'***************************************************
    Valor = Inventory(slot).Valor
End Property

Public Property Get Amount(ByVal slot As Byte) As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's amount
'***************************************************
    If slot = FLAGORO Then
        Amount = UserGLD
    ElseIf slot >= LBound(Inventory) And slot <= UBound(Inventory) Then
        Amount = Inventory(slot).Amount
    End If
End Property

Public Property Get OBJIndex(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's object index
'***************************************************
    OBJIndex = Inventory(slot).OBJIndex
End Property

Public Property Get OBJType(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's object type
'***************************************************
    OBJType = Inventory(slot).OBJType
End Property

Public Property Get ItemName(ByVal slot As Byte) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's name
'***************************************************
    ItemName = Inventory(slot).Name
End Property

Public Property Get Equipped(ByVal slot As Byte) As Boolean
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves True if the item at the given pos is eqiupped
'***************************************************
    Equipped = Inventory(slot).Equipped
End Property

Public Property Get MaxObjs() As Byte
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modify Date: 09/16/09
'Retrieves the capacity of the Inventory
'***************************************************
    MaxObjs = UBound(Inventory)
End Property

Public Sub SetMaxObjs(ByVal MaxObjs As Byte)
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modify Date: 09/16/09
'Set the capacity of the Inventary
'***************************************************
    If UBound(Inventory) = MaxObjs Then Exit Sub
    
    ReDim Preserve Inventory(1 To MaxObjs) As Inventory
End Sub

Public Sub ToggleShowValue(ByVal bShow As Boolean)
'***************************************************
'Author: ZaMa
'Last Modify Date: 11/12
'Indicates if the amount of the items in the inventory is shown
'***************************************************
    ShowValue = bShow
    
    Dim slot As Integer
    ' Update Inventory
    For slot = 1 To Me.MaxObjs
        DrawInventorySlot (slot)
    Next slot
End Sub

Public Sub SetItem(ByVal slot As Byte, ByVal eOBJIndex As Integer, ByVal eAmount As Long, ByVal eEquipped As Byte, _
                        ByVal eGrhIndex As Integer, ByVal eObjType As Integer, ByVal eMaxHit As Integer, ByVal eMinHit As Integer, _
                        ByVal MaxDef As Integer, ByVal MinDef As Integer, ByVal eValor As Single, ByVal eName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/04/06
'Sets all data for a given inventory slot
'***************************************************
    If slot < 1 Or slot > UBound(Inventory) Then Exit Sub
    
    With Inventory(slot)
        .Amount = eAmount
        .MaxDef = MaxDef
        .MinDef = MinDef
        .Equipped = eEquipped
        .GrhIndex = eGrhIndex
        .MaxHit = eMaxHit
        .MinHit = eMinHit
        .Name = eName
        .OBJIndex = eOBJIndex
        .OBJType = eObjType
        .Valor = eValor
    End With
    
    'If InventoryWindow.Visible = False Then InventoryWindow.Visible = True
    
    'Render inventory slot (we don't need to render the whole inventory)
    Call DrawInventorySlot(slot)
    
    Call UpdatelblItem
End Sub

Public Function ClickItem(ByVal x As Long, ByVal Y As Long) As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Selects the item clicked if it's valid and return's it's index
'***************************************************
    Dim TempItem As Long
    Dim temp_x As Long
    Dim temp_y As Long
    If Y >= OFFSETY And x >= OFFSETX Then
        temp_x = (x - OFFSETX) \ ownTilePixelWidth
        temp_y = (Y - OFFSETY) \ ownTilePixelHeight
        
        TempItem = temp_x + (temp_y + InventoryOffset) * ((InventoryWindow.ScaleWidth - OFFSETX) \ ownTilePixelWidth) + 1
        
        'Make sure it's within limits
        If TempItem <= UBound(Inventory) Then
            'Make sure slot isn't empty
            'If Inventory(TempItem).GrhIndex Then
                ClickItem = TempItem
            'Else
            '    ClickItem = 0
            'End If
        End If
    End If
End Function

Private Sub DrawInventory()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Renders the inventory to the given PictureBox
'***************************************************
End Sub

Private Sub DrawInventorySlot(ByVal slot As Integer)
    On Error Resume Next
    Dim SrcRect As RECT
    Dim TempRect As RECT
    Dim destRect As RECT
    Dim SrcRectSlot As RECT
        Dim fileIndex As Integer

    Dim BufferBMP As Long
    With TempRect
        .Bottom = ownTilePixelHeight
        .Right = ownTilePixelWidth
    End With
    
    Dim upperLeftSlot As Integer
    upperLeftSlot = InventoryOffset * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1
    
    'If not in renderable area we exit
    If slot < upperLeftSlot Then Exit Sub
    
    With destRect
        .Top = ((slot - upperLeftSlot) \ ((InventoryWindow.ScaleWidth - OFFSETX) \ ownTilePixelWidth)) * ownTilePixelHeight + OFFSETY
        .Left = ((slot - 1) Mod ((InventoryWindow.ScaleWidth - OFFSETX) \ ownTilePixelWidth)) * ownTilePixelWidth + OFFSETX
        .Bottom = ownTilePixelHeight
        .Right = ownTilePixelWidth
    End With
    
    'Clear the slot area
    'Call InvSurface.BltColorFill(TempRect, InventoryWindow.BackColor)
    If EsMain Then
        Call TransparentBlt(InventoryWindow.hdc, destRect.Left - 1, destRect.Top - 1, destRect.Right + 1, destRect.Bottom + 1, FONDO, destRect.Left - 1, destRect.Top - 1, destRect.Right + 1, destRect.Bottom + 1, vbRed)
    Else
        Call TransparentBlt(InventoryWindow.hdc, destRect.Left - 1, destRect.Top - 1, destRect.Right + 1, destRect.Bottom + 1, FONDO, 0, 0, 32, 32, vbRed)
    End If
    
    If Inventory(slot).GrhIndex Then
     
        'Get source rect
        With SrcRect
            .Left = GrhData(Inventory(slot).GrhIndex).sX
            .Top = GrhData(Inventory(slot).GrhIndex).sY
            .Right = .Left + TilePixelWidth
            .Bottom = .Top + TilePixelHeight
        End With
        
        
        fileIndex = GrhData(Inventory(slot).GrhIndex).FileNum
            
        BufferBMP = Grafico(fileIndex)
        
        
        Call TransparentBlt(InventoryWindow.hdc, destRect.Left, destRect.Top, destRect.Right, destRect.Bottom, BufferBMP, SrcRect.Left, SrcRect.Top, SrcRect.Right, SrcRect.Bottom, ColorTransparente)
            
        
        
        
        If ShowValue Then
        'Render the item grh and the amount
        InventoryWindow.CurrentX = destRect.Left - 2
        InventoryWindow.CurrentY = destRect.Top - 1
        InventoryWindow.ForeColor = 0
        InventoryWindow.Print IIf(Inventory(slot).Amount = 10000, "10k", Inventory(slot).Amount)
        InventoryWindow.CurrentX = destRect.Left - 1
        InventoryWindow.CurrentY = destRect.Top
        InventoryWindow.ForeColor = vbWhite
        InventoryWindow.Print IIf(Inventory(slot).Amount = 10000, "10k", Inventory(slot).Amount)
        End If
        
        'If equipped we render "E"
        If Inventory(slot).Equipped Then
            InventoryWindow.CurrentX = destRect.Left + 23
            InventoryWindow.CurrentY = destRect.Top + 19
            InventoryWindow.ForeColor = 0
            InventoryWindow.Print "+"
            InventoryWindow.CurrentX = destRect.Left + 24
            InventoryWindow.CurrentY = destRect.Top + 20
            InventoryWindow.ForeColor = vbYellow
            InventoryWindow.Print "+"
        End If
        
        'Render a box around the selected item
        If InvSelectedItem = slot Then
            InventoryWindow.ForeColor = RGB(211, 21, 20)
            InventoryWindow.FillStyle = 1
            InventoryWindow.DrawStyle = 3
            
            InventoryWindow.Line (destRect.Left, destRect.Top)-(destRect.Left + TilePixelWidth - 2, destRect.Top + TilePixelHeight - 2), , B
        End If
        
        
    End If

    Call InventoryWindow.Refresh
End Sub

    Public Sub Initialize(ByRef InvPic As PictureBox, mOffX As Integer, mOffY As Integer, ByVal MaxObjs As Byte, bEsMain As Boolean, _
                            Optional ByVal FontSize As Integer = 7, Optional ByVal TileWidth As Integer = 32, _
                            Optional ByVal TileHeight As Integer = 32, Optional ByVal startX As Integer = 0, _
                            Optional ByVal startY As Integer = 0, Optional ByVal bImgContainer As Boolean = True, _
                            Optional ByVal bShowText As Boolean = True)

'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 03/12/09
'Sets the reference to the picture box on which to render the inventory
'03/12/09: I added the optionals args FontSize, TileWidth, TileHeight, startX and startY. (Budi)
'***************************************************
    Set InventoryWindow = InvPic
    
    'Make sure auto-redraw is set to true
    InventoryWindow.AutoRedraw = True
    
    'Set apropiate scale (pixel)
    InventoryWindow.ScaleMode = 3
    
    ReDim Inventory(1 To MaxObjs) As Inventory
    
    ownTilePixelHeight = TileHeight
    ownTilePixelWidth = TileWidth
    startdX = startX
    startdY = startY
    
    ImgContainer = bImgContainer
    ShowValue = bShowText
   
'TODO : Fonts should be in a separate class / collection
    Dim font As New StdFont
    Dim Ifnt As IFont
    
    font.Name = "Verdana"
    font.bold = True
    font.italic = False
    font.Size = FontSize
    font.Underline = False
    font.Strikethrough = False
    
    Set Ifnt = font
    
    
    If ImgContainer Then _
        InvSelectedItem = ClickItem(1, 1)   'If there is anything there we select the top left item
    
    InventoryWindow.Cls
    
    FONDO = CreateCompatibleDC(InventoryWindow.hdc)

    EsMain = bEsMain
    If EsMain Then
        Call SelectObject(FONDO, LoadPictureEX("VentanaPrincipalInv.jpg").handle)
    Else
        Call SelectObject(FONDO, Grafico(1))
        Dim R As pvRECT
        Dim hBr As Long, TmpBMP As Long
        R.Bottom = 32
        R.Right = 32
        TmpBMP = CreateCompatibleBitmap(InventoryWindow.hdc, 32, 32)
        SelectObject FONDO, TmpBMP
        
        hBr = CreateSolidBrush(&H0)
        FillRect FONDO, R, hBr
        DeleteObject hBr
    End If
    OFFSETX = mOffX
    OFFSETY = mOffY

End Sub

Private Sub InventoryWindow_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Implements the mouse up event of the inventory picture box
'Check outs which item was clicked
'***************************************************
    'Store previously selected item
    Dim prevSelItem As Long
    
    'Exit if it got outside the control's area
    If x < OFFSETX Or Y < OFFSETY Or x > InventoryWindow.Width Or Y > InventoryWindow.Height Then _
        Exit Sub
    
    prevSelItem = InvSelectedItem

    'Get the currently clickced item
    InvSelectedItem = ClickItem(CInt(x), CInt(Y))
    
    'Update needed inventory slots
    If prevSelItem <> InvSelectedItem Then
        If prevSelItem <> 0 And prevSelItem <> FLAGORO Then Call DrawInventorySlot(prevSelItem)
        If InvSelectedItem Then Call DrawInventorySlot(InvSelectedItem)
    End If
    
    If EsMain And InvSelectedItem > 0 Then
        UpdatelblItem
    End If
End Sub
Private Sub UpdatelblItem()
If InvSelectedItem > 0 And InvSelectedItem <= 20 Then
frmMain.lblItem.Caption = Inventory(InvSelectedItem).Name
If Inventory(InvSelectedItem).MaxHit > 0 Then
    frmMain.lblItem.Caption = frmMain.lblItem.Caption & " [" & Inventory(InvSelectedItem).MinHit & "/" & Inventory(InvSelectedItem).MaxHit & "]"
ElseIf Inventory(InvSelectedItem).MaxDef > 0 Then
    frmMain.lblItem.Caption = frmMain.lblItem.Caption & " [" & Inventory(InvSelectedItem).MinDef & "/" & Inventory(InvSelectedItem).MaxDef & "]"
End If
End If
End Sub


Private Sub InventoryWindow_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Implements the mouse move event of the inventory picture box
'Displays a ToolTip of the item under the mouse
'***************************************************
    Dim temp_x As Integer
    Dim temp_y As Integer
    Dim TempItem As Integer
    
    'Exit if it got outside the control's area
    If x < 0 Or Y < 0 Or x > InventoryWindow.Width Or Y > InventoryWindow.Height Then _
        Exit Sub
    
    temp_x = x \ ownTilePixelWidth
    temp_y = Y \ ownTilePixelHeight
    
    TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1
    
    If TempItem <= UBound(Inventory) Then
        InventoryWindow.ToolTipText = Inventory(TempItem).Name
    End If
End Sub

Public Function Grafico(ByVal fileIndex As Integer) As Long
Dim buffer As Long
Dim bmpInfo As BITMAPINFO
Dim BufferBMP As Long
Dim Data() As Byte
If Not Graficos.Exists("G" & fileIndex) Then
    'get Bitmap
    Call Get_Bitmap(DirRecursos & "Graphics.AO", CStr(fileIndex) & ".BMP", bmpInfo, Data)

    BufferBMP = CreateCompatibleDC(InventoryWindow.hdc)
    buffer = CreateCompatibleBitmap(InventoryWindow.hdc, bmpInfo.bmiHeader.biWidth, bmpInfo.bmiHeader.biHeight)
    SelectObject BufferBMP, buffer
    Call SetDIBitsToDevice(BufferBMP, 0, 0, bmpInfo.bmiHeader.biWidth, bmpInfo.bmiHeader.biHeight, 0, 0, 0, bmpInfo.bmiHeader.biHeight, Data(0), bmpInfo, DIB_RGB_COLORS)
    Graficos.Add "G" & fileIndex, BufferBMP
End If
Grafico = Graficos("G" & fileIndex)
End Function

Public Sub Intercambiar(ByVal NewPosItem As Integer)
    Dim AntPosItem As Integer
    Dim tmpInv As Inventory
    AntPosItem = InvSelectedItem
    If NewPosItem > 0 And NewPosItem <> InvSelectedItem Then
        tmpInv = Inventory(InvSelectedItem)
        Inventory(InvSelectedItem) = Inventory(NewPosItem)
        Inventory(NewPosItem) = tmpInv
        InvSelectedItem = NewPosItem
        Call DrawInventorySlot(AntPosItem)
        Call DrawInventorySlot(NewPosItem)
    End If
End Sub
