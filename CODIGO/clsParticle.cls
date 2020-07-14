VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParticle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mvarsngX As Single
Private mvarsngY As Single
Private mvarsngSize As Single
Private mvarsngXSpeed As Single
Private mvarsngYSpeed As Single
Private mvarsngXAccel As Single
Private mvarsngYAccel As Single
Private mvarsngR As Single
Private mvarsngG As Single
Private mvarsngB As Single
Private mvarsngA As Single
Private mvarsngAlphaDecay As Single

Public Property Let sngAlphaDecay(ByVal vData As Single)
    mvarsngAlphaDecay = vData
End Property

Public Property Get sngAlphaDecay() As Single
    sngAlphaDecay = mvarsngAlphaDecay
End Property

Public Property Let sngA(ByVal vData As Single)
    mvarsngA = vData
End Property

Public Property Get sngA() As Single
    sngA = mvarsngA
End Property

Public Property Let sngB(ByVal vData As Single)
    mvarsngB = vData
End Property

Public Property Get sngB() As Single
    sngB = mvarsngB
End Property

Public Property Let sngG(ByVal vData As Single)
    mvarsngG = vData
End Property

Public Property Get sngG() As Single
    sngG = mvarsngG
End Property

Public Property Let sngR(ByVal vData As Single)
    mvarsngR = vData
End Property

Public Property Get sngR() As Single
    sngR = mvarsngR
End Property

Public Property Let sngYAccel(ByVal vData As Single)
    mvarsngYAccel = vData
End Property

Public Property Get sngYAccel() As Single
    sngYAccel = mvarsngYAccel
End Property

Public Property Let sngXAccel(ByVal vData As Single)
    mvarsngXAccel = vData
End Property

Public Property Get sngXAccel() As Single
    sngXAccel = mvarsngXAccel
End Property

Public Property Let sngYSpeed(ByVal vData As Single)
    mvarsngYSpeed = vData
End Property

Public Property Get sngYSpeed() As Single
    sngYSpeed = mvarsngYSpeed
End Property

Public Property Let sngXSpeed(ByVal vData As Single)
    mvarsngXSpeed = vData
End Property

Public Property Get sngXSpeed() As Single
    sngXSpeed = mvarsngXSpeed
End Property

Public Property Let sngSize(ByVal vData As Single)
    mvarsngSize = vData
End Property

Public Property Get sngSize() As Single
    sngSize = mvarsngSize
End Property

Public Property Let sngY(ByVal vData As Single)
    mvarsngY = vData
End Property

Public Property Get sngY() As Single
    sngY = mvarsngY
End Property

Public Property Let sngX(ByVal vData As Single)
    mvarsngX = vData
End Property

Public Property Get sngX() As Single
    sngX = mvarsngX
End Property

Public Sub ResetColor(sngRed As Single, sngGreen As Single, sngBlue As Single, sngAlpha As Single, sngDecay As Single)
    '//Reset color to the new values
    sngR = sngRed
    sngG = sngGreen
    sngB = sngBlue
    sngA = sngAlpha
    sngAlphaDecay = sngDecay
End Sub

Public Sub ResetIt(X As Single, Y As Single, XSpeed As Single, YSpeed As Single, XAcc As Single, YAcc As Single, sngResetSize As Single)
    sngX = X
    sngY = Y
    sngXSpeed = XSpeed
    sngYSpeed = YSpeed
    sngXAccel = XAcc
    sngYAccel = YAcc
    sngSize = sngResetSize
End Sub

Public Sub UpdateParticle(sngTime As Single)
    sngX = sngX + sngXSpeed * sngTime
    sngY = sngY + sngYSpeed * sngTime
    
    sngXSpeed = sngXSpeed + sngXAccel * sngTime
    sngYSpeed = sngYSpeed + sngYAccel * sngTime
    
    sngA = sngA - sngAlphaDecay * sngTime
End Sub

