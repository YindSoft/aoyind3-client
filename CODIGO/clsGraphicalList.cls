VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGraphicalList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Por Javier Podavini (El Yind)
'09/07/2009

Private ListPic As PictureBox
Private Lista() As String
Private p_ListIndex As Integer
Private p_ListCount As Integer
Private p_Visible As Boolean
Private p_ForeColor As Long
Private p_Pixel_Alto As Integer
Private p_Scroll As Integer
Private p_OldY As Integer
Private Sub Class_Initialize()
p_Visible = True
Clear
End Sub

Public Sub Initialize(ByRef LPic As PictureBox, ByVal ForeColor As Long)
    Set ListPic = LPic
    
    p_ForeColor = ForeColor
    p_Pixel_Alto = 13
    
    ListPic.ForeColor = p_ForeColor
    ListPic.Cls
    
End Sub

Public Property Get List(ByVal Index As Integer) As String
Attribute List.VB_UserMemId = 0
    List = Lista(Index)
End Property
Public Property Let List(ByVal Index As Integer, ByVal Valor As String)
    Lista(Index) = Valor
    Dibujar
End Property

Public Property Get Visible() As Integer
    Visible = p_Visible
End Property
Public Property Let Visible(ByVal Valor As Integer)
    p_Visible = Valor
    ListPic.Visible = p_Visible
End Property

Public Property Get ListIndex() As Integer
    ListIndex = p_ListIndex
End Property
Public Property Let ListIndex(ByVal Valor As Integer)
    p_ListIndex = Valor
    If p_ListIndex >= p_ListCount Then p_ListIndex = p_ListCount - 1
    Dibujar
End Property

Public Property Get OldY() As Integer
    OldY = p_OldY
End Property
Public Property Let OldY(ByVal Valor As Integer)
    p_OldY = Valor
End Property

Public Property Get Pixel_Alto() As Integer
    Pixel_Alto = p_Pixel_Alto
End Property
Public Property Let Pixel_Alto(ByVal Valor As Integer)
    p_Pixel_Alto = Valor
End Property

Public Property Get Scroll() As Integer
    Scroll = p_Scroll
End Property
Public Property Let Scroll(ByVal Valor As Integer)
    p_Scroll = Valor
    Dibujar
End Property

Public Property Get ListCount() As Integer
    ListCount = p_ListCount
End Property

Public Sub AddItem(item As String)
ReDim Preserve Lista(p_ListCount)
Lista(p_ListCount) = item
p_ListCount = p_ListCount + 1
If p_ListCount = 1 Then p_ListIndex = 0
Dibujar
End Sub
Public Sub Clear()
ReDim Lista(0)
p_ListCount = 0
p_ListIndex = -1
p_Scroll = 0
If Not ListPic Is Nothing Then Dibujar
End Sub
Private Sub Dibujar()
Dim i As Integer
Dim PosY As Integer
ListPic.Cls
For i = p_Scroll To p_Scroll + Int(ListPic.Height / p_Pixel_Alto)
    If i < p_ListCount Then
        PosY = (i - p_Scroll) * p_Pixel_Alto + 4

        If i = p_ListIndex Then
            ListPic.ForeColor = RGB(211, 120, 20)
            ListPic.Line (1, PosY - 1)-(ListPic.Width - 2, PosY + p_Pixel_Alto + 1), , B
            ListPic.ForeColor = RGB(130, 130, 130)
        Else
            ListPic.ForeColor = p_ForeColor
        End If
        ListPic.CurrentX = 4
        ListPic.CurrentY = PosY
       
        ListPic.Print Lista(i)
    End If
Next i
End Sub
