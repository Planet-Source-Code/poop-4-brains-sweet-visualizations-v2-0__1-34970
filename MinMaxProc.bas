Attribute VB_Name = "MinMaxProc"
'----------------------------------------------------------------------
' This is a dopey window proceedure that restricts the size of a
' window.
'----------------------------------------------------------------------
' Murphy McCauley (MurphyMc@Concentric.NET) 08/06/99
'----------------------------------------------------------------------

Option Explicit

Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal HWnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal HWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal HWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Proc As Long
Public Const GWL_WNDPROC = (-4)

Private Const WM_GETMINMAXINFO = &H24

Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef Destination As Any, _
    ByRef Source As Any, _
    ByVal ByteLength As Long)

Private Type POINTAPI
        x As Long
        y As Long
End Type

Private Type MINMAXINFO
        ptReserved As POINTAPI
        ptMaxSize As POINTAPI
        ptMaxPosition As POINTAPI
        ptMinTrackSize As POINTAPI
        ptMaxTrackSize As POINTAPI
End Type

Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Function WindowProc(ByVal HWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Static MinMax As MINMAXINFO
    
    Select Case Msg
        Case WM_GETMINMAXINFO
            Call MoveMemory(MinMax, ByVal lParam, Len(MinMax))
            MinMax.ptMinTrackSize.x = Base.MinWidth
            MinMax.ptMinTrackSize.y = Base.MinHeight
            Call MoveMemory(ByVal lParam, MinMax, Len(MinMax))
        Case Else
            WindowProc = CallWindowProc(Proc, HWnd, Msg, wParam, lParam)
    End Select
End Function
