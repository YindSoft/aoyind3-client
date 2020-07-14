Attribute VB_Name = "modVentanas"
Option Explicit
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal color As Long, ByVal bAlpha As Byte, ByVal Alpha As Long) As Boolean
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function OleLoadPicture Lib "olepro32" (pStream As Any, ByVal lSize As Long, ByVal fRunmode As Long, riid As Any, ppvObj As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpsz As Any, pclsid As Any) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED As Long = &H80000
Private Const LWA_ALPHA As Long = &H2

Public Const NTRANS_GENERAL As Integer = 200

Public Sub SetTranslucent(ThehWnd As Long, nTrans As Integer)
On Error GoTo ErrorRtn

   Dim attrib As Long

   'put current GWL_EXSTYLE in attrib
   attrib = GetWindowLong(ThehWnd, GWL_EXSTYLE)

   'change GWL_EXSTYLE to WS_EX_LAYERED - makes a window layered
   SetWindowLong ThehWnd, GWL_EXSTYLE, attrib Or WS_EX_LAYERED

   'Make transparent (RGB value does not have any effect at this
   'time, will in Part 2 of this article)
   SetLayeredWindowAttributes ThehWnd, RGB(0, 0, 0), nTrans, _
                                       LWA_ALPHA
   Exit Sub

ErrorRtn:
MsgBox Err.Description & " Source : " & Err.Source

End Sub

Public Sub MoverVentana(hwnd As Long)
    ReleaseCapture
    SendMessage hwnd, &HA1, 2, 0&
End Sub

Public Sub MessageBox(ByVal Message As String, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal Title As String = "")
frmMensaje.msg.Caption = Message
frmMensaje.Show
End Sub
Public Function LoadPictureEX(ByVal FileName As String) As IPicture
If FileName = "" Then
    Set LoadPictureEX = Nothing
Else
    Dim b() As Byte
    Call Get_File_Data(DirRecursos & "Interface.AO", FileName, b)
    Set LoadPictureEX = PictureFromByteStream(b)
End If
End Function
Public Function PictureFromByteStream(ByRef b() As Byte) As IPicture
    Dim LowerBound As Long
    Dim ByteCount  As Long
    Dim hMem  As Long
    Dim lpMem  As Long
    Dim IID_IPicture(15)
    Dim istm As stdole.IUnknown

    On Error GoTo Err_Init
    If UBound(b, 1) < 0 Then
        Exit Function
    End If
    
    LowerBound = LBound(b)
    ByteCount = (UBound(b) - LowerBound) + 1
    hMem = GlobalAlloc(&H2, ByteCount)
    If hMem <> 0 Then
        lpMem = GlobalLock(hMem)
        If lpMem <> 0 Then
            MoveMemory ByVal lpMem, b(LowerBound), ByteCount
            Call GlobalUnlock(hMem)
            If CreateStreamOnHGlobal(hMem, 1, istm) = 0 Then
                If CLSIDFromString(StrPtr("{7BF80980-BF32-101A-8BBB-00AA00300CAB}"), IID_IPicture(0)) = 0 Then
                  Call OleLoadPicture(ByVal ObjPtr(istm), ByteCount, 0, IID_IPicture(0), PictureFromByteStream)
                End If
            End If
        End If
    End If
    
    Exit Function
    
Err_Init:

        MsgBox Err.Number & " - " & Err.Description

End Function
