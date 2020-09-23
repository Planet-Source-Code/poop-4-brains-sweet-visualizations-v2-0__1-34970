Attribute VB_Name = "modCaps"

Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public CapVal(255) As Long 'the cap's value
Public CapSp(255) As Long

Public intX As Integer, intY As Integer
Public intI As Integer, intJ As Integer

Public MoveM

Public LsrI As Long
Public LsrIStp As Long

Public LsrV(257) As Long
Public LsrC(257) As Long
Public LsrX(257) As Long

Public ShpX(255) As Double
Public ShpY(255) As Double
Public ShpT(255) As Integer
Public ShpC(255) As Integer

Type Color
r As Integer
g As Integer
b As Integer
End Type

Public OldVal(520, 1 To 10) As Byte
Public OldColor(1 To 10) As Color

Public OldI As Long

Public Const MaxFade = 10

Public DreamClr As Integer
Public DreamCnt As Double

Public FPS, TFPS

Public Function Rand(L, U) As Long
    Dim I As Long, U2, L2
    If U < 0 Then
    U2 = U
    Else
    U2 = -U
    End If
    
    If L < 0 Then
    L2 = L
    Else
    L2 = -L
    End If
    
    For I = L To U
        If Int(Rnd * (U2 - L2)) = Int(Rnd * (U2 - L2 + 1)) Then Rand = I: Exit Function
    Next I
End Function

Function ArrayAdd(Nums() As Byte, Step As Integer)
Dim I As Long
For I = LBound(Nums()) To UBound(Nums()) Step Step
ArrayAdd = ArrayAdd + Val(Nums(I))
Next I
End Function

Function MassAdd(ParamArray Nums() As Variant) As Double
Dim I As Integer
For I = 0 To UBound(Nums())
MassAdd = MassAdd + Val(Nums(I))
Next I
End Function

Function MakeRandMove(Power, Index)
Power = Power * 0.6
Select Case Int(Rnd * 6)
Case 0: MakeMove Power, 0, Index
Case 1: MakeMove Power, Power, Index
Case 2: MakeMove 0, Power, Index
Case 3: MakeMove -Power, 0, Index
Case 4: MakeMove -Power, -Power, Index
Case 5: MakeMove 0, -Power, Index
End Select
End Function

Function MakeMove(XS, YS, Index)
ShpX(Index) = ShpY(Index) + XS
ShpY(Index) = ShpY(Index) + YS
End Function
