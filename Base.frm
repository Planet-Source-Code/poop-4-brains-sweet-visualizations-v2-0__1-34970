VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Base 
   Caption         =   "Sweet Visualizations v2.0"
   ClientHeight    =   6795
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10455
   Icon            =   "Base.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   453
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   697
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CMN 
      Left            =   1440
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open Song"
      Filter          =   "MP3 Files|*.mp3*"
   End
   Begin VB.Timer tmrFPS 
      Interval        =   1000
      Left            =   2880
      Top             =   1200
   End
   Begin VB.PictureBox grad 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   4500
      Left            =   6840
      Picture         =   "Base.frx":5C12
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   4
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.PictureBox Board 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2160
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   4
      Top             =   1560
      Width           =   255
   End
   Begin VB.Frame Stuff 
      BorderStyle     =   0  'None
      Height          =   336
      Left            =   720
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   3360
      Begin VB.CommandButton StartButton 
         Caption         =   "&Start"
         Height          =   336
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   804
      End
      Begin VB.CommandButton StopButton 
         Caption         =   "S&top"
         Enabled         =   0   'False
         Height          =   336
         Left            =   864
         TabIndex        =   2
         Top             =   0
         Width           =   804
      End
   End
   Begin VB.ComboBox cmbV 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      ItemData        =   "Base.frx":6A64
      Left            =   1560
      List            =   "Base.frx":6AAA
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   0
      Width           =   6735
   End
   Begin VB.ComboBox DevicesBox 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   315
   End
   Begin MediaPlayerCtl.MediaPlayer med 
      Height          =   645
      Left            =   720
      TabIndex        =   8
      Top             =   2640
      Width           =   9135
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   -1  'True
      Balance         =   10
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   0
      WindowlessVideo =   0   'False
   End
   Begin VB.Label lblX 
      BackColor       =   &H00000000&
      Caption         =   "Normal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Begin VB.Menu mnuFull 
         Caption         =   "Full Screen"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuChoose 
      Caption         =   "Choose Song"
   End
End
Attribute VB_Name = "Base"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------
' Deeth Stereo Oscilloscope v1.0
' A simple oscilloscope application -- now in <<stereo>>
'----------------------------------------------------------------------
' Opens a waveform audio device for 8-bit 11kHz input, and plots the
' waveform to a window.  Can only be resized to a certain minimum
' size defined by the Shape box.
'----------------------------------------------------------------------
' It would be good to make this use the same double-buffering
' scheme as the Spectrum Analyzer.
'----------------------------------------------------------------------
' Murphy McCauley (MurphyMc@Concentric.NET) 08/12/99
'----------------------------------------------------------------------

Option Explicit

Private DevHandle As Long
Private InData(0 To 511) As Byte
Private InOldD(0 To 511) As Byte

Private Inited As Boolean
Public MinHeight As Long, MinWidth As Long

Private Type WaveFormatEx
    FormatTag As Integer
    Channels As Integer
    SamplesPerSec As Long
    AvgBytesPerSec As Long
    BlockAlign As Integer
    BitsPerSample As Integer
    ExtraDataSize As Integer
End Type

Private Type WaveHdr
    lpData As Long
    dwBufferLength As Long
    dwBytesRecorded As Long
    dwUser As Long
    dwFlags As Long
    dwLoops As Long
    lpNext As Long 'wavehdr_tag
    Reserved As Long
End Type

Private Type WaveInCaps
    ManufacturerID As Integer      'wMid
    ProductID As Integer       'wPid
    DriverVersion As Long       'MMVERSIONS vDriverVersion
    ProductName(1 To 32) As Byte 'szPname[MAXPNAMELEN]
    Formats As Long
    Channels As Integer
    Reserved As Integer
End Type

Private Const WAVE_INVALIDFORMAT = &H0&                 '/* invalid format */
Private Const WAVE_FORMAT_1M08 = &H1&                   '/* 11.025 kHz, Mono,   8-bit
Private Const WAVE_FORMAT_1S08 = &H2&                   '/* 11.025 kHz, Stereo, 8-bit
Private Const WAVE_FORMAT_1M16 = &H4&                   '/* 11.025 kHz, Mono,   16-bit
Private Const WAVE_FORMAT_1S16 = &H8&                   '/* 11.025 kHz, Stereo, 16-bit
Private Const WAVE_FORMAT_2M08 = &H10&                  '/* 22.05  kHz, Mono,   8-bit
Private Const WAVE_FORMAT_2S08 = &H20&                  '/* 22.05  kHz, Stereo, 8-bit
Private Const WAVE_FORMAT_2M16 = &H40&                  '/* 22.05  kHz, Mono,   16-bit
Private Const WAVE_FORMAT_2S16 = &H80&                  '/* 22.05  kHz, Stereo, 16-bit
Private Const WAVE_FORMAT_4M08 = &H100&                 '/* 44.1   kHz, Mono,   8-bit
Private Const WAVE_FORMAT_4S08 = &H200&                 '/* 44.1   kHz, Stereo, 8-bit
Private Const WAVE_FORMAT_4M16 = &H400&                 '/* 44.1   kHz, Mono,   16-bit
Private Const WAVE_FORMAT_4S16 = &H800&                 '/* 44.1   kHz, Stereo, 16-bit

Private Const WAVE_FORMAT_PCM = 1

Private Const WHDR_DONE = &H1&              '/* done bit */
Private Const WHDR_PREPARED = &H2&          '/* set if this header has been prepared */
Private Const WHDR_BEGINLOOP = &H4&         '/* loop start block */
Private Const WHDR_ENDLOOP = &H8&           '/* loop end block */
Private Const WHDR_INQUEUE = &H10&          '/* reserved for driver */

Private Const WIM_OPEN = &H3BE
Private Const WIM_CLOSE = &H3BF
Private Const WIM_DATA = &H3C0

Private Declare Function waveInAddBuffer Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long
Private Declare Function waveInPrepareHeader Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long
Private Declare Function waveInUnprepareHeader Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long

Private Declare Function waveInGetNumDevs Lib "winmm" () As Long
Private Declare Function waveInGetDevCaps Lib "winmm" Alias "waveInGetDevCapsA" (ByVal uDeviceID As Long, ByVal WaveInCapsPointer As Long, ByVal WaveInCapsStructSize As Long) As Long

Private Declare Function waveInOpen Lib "winmm" (WaveDeviceInputHandle As Long, ByVal WhichDevice As Long, ByVal WaveFormatExPointer As Long, ByVal CallBack As Long, ByVal CallBackInstance As Long, ByVal Flags As Long) As Long
Private Declare Function waveInClose Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long

Private Declare Function waveInStart Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long
Private Declare Function waveInReset Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long
Private Declare Function waveInStop Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long

'normal visuals
Const vBar = 0
Const vCircle = 1
Const vColors = 2
Const vExplo = 3
Const vLines = 4
Const vScope = 5
Const vBackC = 6
Const vGradBars = 7
Const vIce = 8
Const vImp = 9
Const vWalls = 12
Const vShapes = 13

'lasers
Const vLaser = 10
Const vSLaser = 11

'dreams
Const vGDream = 14
Const vYDream = 15
Const vADream = 16
Const vPDream = 17
Const vGrDream = 18
Const vBDream = 19
Const vRDream = 20
Const vRainDream = 21

Const PI = 3.14

Dim VMode As Long

Sub InitDevices()
    Dim Caps As WaveInCaps, Which As Long
    DevicesBox.Clear
    For Which = 0 To waveInGetNumDevs - 1
        Call waveInGetDevCaps(Which, VarPtr(Caps), Len(Caps))
        'If Caps.Formats And WAVE_FORMAT_1M08 Then
        If Caps.Formats And WAVE_FORMAT_1S08 Then 'Now is 1S08 -- Check for devices that can do stereo 8-bit 11kHz
            Call DevicesBox.AddItem(StrConv(Caps.ProductName, vbUnicode), Which)
        End If
    Next
    If DevicesBox.ListCount = 0 Then
        MsgBox "You have no audio input devices!", vbCritical, "Ack!"
        End
    End If
    DevicesBox.ListIndex = 0
End Sub

Private Sub Form_Load()
Dim I As Long
For I = 0 To 255
CapSp(I) = 1
Next I
cmbV.ListIndex = 0
InitDevices
Me.Visible = True
StartButton_Click
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then Exit Sub
cmbV.Width = Me.ScaleWidth
cmbV.Left = 0
Board.Top = cmbV.Height
Board.Left = 0
Board.Width = Me.ScaleWidth
Board.Height = Me.ScaleHeight - Board.Top - med.Height
med.Width = Me.ScaleWidth
med.Top = Me.ScaleHeight - med.Height
med.Left = 0
If lblX.Visible = True Then
cmbV.Width = Me.ScaleWidth - lblX.Width
cmbV.Left = lblX.Width
lblX.Height = cmbV.Height
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If DevHandle <> 0 Then
        Call DoStop
    End If
    Board.Cls
    End
End Sub

Private Sub lblX_Click()
lblX.Visible = False
Me.WindowState = vbNormal
Me.BorderStyle = 2
End Sub

Private Sub mnuAbout_Click()
MsgBox App.Comments, vbInformation, "About Sweet Visualizations"
End Sub

Private Sub mnuChoose_Click()
Dim path
CMN.ShowOpen
path = CMN.FileName
If path = "" Then Exit Sub
med.FileName = path
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuFull_Click()
Me.BorderStyle = 0
lblX.Visible = True
Me.WindowState = vbMaximized
End Sub

Private Sub mnuPath_Click()

End Sub

Private Sub mnuVis_Click()

End Sub

Private Sub StartButton_Click()
    Dim x As Long
    Randomize
    For x = 0 To 255 Step 5
    ShpX(x) = Int(Rnd * Board.ScaleWidth)
    ShpY(x) = Int(Rnd * Board.ScaleHeight)
    ShpT(x) = Int(Rnd * 3)
    ShpC(x) = Int(Rnd * 3)
    Next x
    
    Static WaveFormat As WaveFormatEx
    With WaveFormat
        .FormatTag = WAVE_FORMAT_PCM
        .Channels = 2 'Two channels -- left and right
        .SamplesPerSec = 11025 '11khz
        .BitsPerSample = 8
        .BlockAlign = (.Channels * .BitsPerSample) \ 8
        .AvgBytesPerSec = .BlockAlign * .SamplesPerSec
        .ExtraDataSize = 0
    End With
    
    Debug.Print "waveInOpen:"; waveInOpen(DevHandle, DevicesBox.ListIndex, VarPtr(WaveFormat), 0, 0, 0)
    
    If DevHandle = 0 Then
        Call MsgBox("Wave input device didn't open!", vbExclamation, "Ack!")
        Exit Sub
    End If
    Debug.Print " "; DevHandle
    Call waveInStart(DevHandle)
    
    Inited = True
       
    StopButton.Enabled = True
    StartButton.Enabled = False
    
    Call Visualize
End Sub


Private Sub StopButton_Click()
    Call DoStop
End Sub


Private Sub DoStop()
    Call waveInReset(DevHandle)
    Call waveInClose(DevHandle)
    DevHandle = 0
End Sub


Private Sub Visualize()
    Static Wave As WaveHdr
    
    Wave.lpData = VarPtr(InData(0))
    Wave.dwBufferLength = 512 'This is now 512 so there's still 256 samples per channel
    Wave.dwFlags = 0
    
    Do
    
        Call waveInPrepareHeader(DevHandle, VarPtr(Wave), Len(Wave))
        Call waveInAddBuffer(DevHandle, VarPtr(Wave), Len(Wave))
    
        Do
            'Nothing -- we're waiting for the audio driver to mark
            'this wave chunk as done.
        Loop Until ((Wave.dwFlags And WHDR_DONE) = WHDR_DONE) Or DevHandle = 0
        
        Call waveInUnprepareHeader(DevHandle, VarPtr(Wave), Len(Wave))
        
        If DevHandle = 0 Then
            'The device has closed...
            Exit Do
        End If
        
        Call DrawData
        
        DoEvents
    Loop While DevHandle <> 0 'While the audio device is open

End Sub

Function DrawData()
Static x As Long, g

If cmbV.ListIndex < vGDream And cmbV.ListIndex < vRainDream Then Board.Cls
 
Select Case cmbV.ListIndex
Case vBar 'reg bars
    
    'right
    For x = 0 To 255
        Board.Line (0, x * 5)-(InData(x * 2), x * 5 + 3), vbGreen, BF
    Next x
    
    'left
    For x = 0 To 255
        Board.Line (Board.ScaleWidth, x * 5)-(Board.ScaleWidth - InData(x * 2), x * 5 + 3), vbRed, BF
    Next x

Case vCircle 'circle scope
    
    For x = 0 To 255
        Board.Circle (Board.ScaleWidth \ 2, x * (Board.ScaleHeight \ 126)), InData(x * 2) \ 2, vbBlue
    Next x

Case vColors 'colored squares
    
    Dim Width
    
    For x = 0 To 255 Step 5
        Width = InData(x * 2) * 2
        Board.Line (Board.ScaleWidth \ 2 - Width \ 2, Board.ScaleHeight \ 2 - Width \ 2)-(Board.ScaleWidth \ 2 + Width \ 2, Board.ScaleHeight \ 2 + Width \ 2), RGB(x, x, x), BF
    Next x
    
Case vExplo 'explo
    
    For x = 0 To 255
        Board.Circle (Board.ScaleWidth \ 2, Board.ScaleHeight \ 2), InData(x * 2), RGB(x, x, x)
    Next x
    
Case vLines 'lines
    
    For x = 0 To 254
        Board.Line (Board.ScaleWidth, Board.ScaleHeight \ 2)-(Board.ScaleWidth \ 2, InData(x * 2 + 2)), RGB(x, 0, 0)
        Board.Line (0, Board.ScaleHeight \ 2)-(Board.ScaleWidth \ 2, InData(x * 2)), RGB(x, 0, 0)
    Next x
    
Case vScope 'scope
    
    Dim Stp As Long, dX As Long
    Stp = Board.ScaleWidth \ 255

    'right
    For x = 0 To 255
        dX = x * Stp
        Board.Line (Board.CurrentX, Board.CurrentY)-(dX * 2, InData(x * 2)), vbBlue, BF
    Next x
    
    Board.CurrentX = 0
    Board.CurrentY = Board.ScaleWidth
    
    'left
    For x = 0 To 255
        dX = x * Stp
        Board.Line (Board.CurrentX, Board.CurrentY)-(dX * 2, InData(x * 2 + 1)), vbRed, BF
    Next x
    
Case vBackC 'climate colors
    
    Dim Total As Double, Avg As Integer
    
    For x = 0 To 255
    Total = Total + InData(x)
    Next x
    Avg = Total / 255
    Board.Line (0, 0)-(Board.ScaleWidth, Board.ScaleHeight), RGB(Avg, 0, 0), BF
    Total = 0
    Avg = 0
Case vGradBars 'gradient bars
    
    For x = 0 To 255
    CapVal(x) = CapVal(x) + CapSp(x)
    CapSp(x) = CapSp(x) - 1
    If InData(x * 2) > CapVal(x) Then CapVal(x) = InData(x * 2) + 10: CapSp(x) = -5
    BitBlt Board.hDC, x * 5, Board.ScaleHeight - InData(x * 2), 4, InData(x * 2), grad.hDC, 0, grad.ScaleHeight - InData(x * 2), vbSrcCopy
    BitBlt Board.hDC, x * 5, Board.ScaleHeight - CapVal(x), 4, 3, grad.hDC, 0, grad.ScaleHeight - CapVal(x), vbSrcCopy
    Next x
    
Case vIce


    Dim N As Long, Color
    
    Stp = Board.ScaleWidth \ (254 / 2)
    Board.CurrentY = Board.ScaleHeight / 2
    
    For N = 1 To MaxFade
        For x = 0 To (254 / 2)
            dX = x * Stp * 1.5
            Color = RGB(OldColor(N).r, OldColor(N).g, OldColor(N).b)
            Board.Line ((x + 2) * Stp * 1.5, OldVal(x * 4 + 4, N))-(dX, OldVal(x * 4, N)), Color
        Next x
        OldColor(N).r = OTrim(OldColor(N).r - 35)
        OldColor(N).g = OTrim(OldColor(N).g - 20)
        OldColor(N).b = OTrim(OldColor(N).b - 20)
    Next N

        OldI = OldI + 1: If OldI > 5 Then OldI = 1
        OldColor(OldI).r = 100
        OldColor(OldI).g = 100
        OldColor(OldI).b = 255
        For N = 0 To (254 / 2)
            OldVal(N * 4, OldI) = InData(N * 4)
        Next N
        
    Board.CurrentY = Board.ScaleHeight / 2
    Board.CurrentX = 0
    
    For x = 0 To (254 / 2)
        dX = x * Stp * 1.5
        Board.Line (Board.CurrentX, Board.CurrentY)-(dX, InData(x * 4)), vbBlue
    Next x
    
Case vImp

Dim mak As Double, dstX, dstY, cr As Integer, cg As Integer, cb As Integer, ang As Long, olddstx, olddsty
mak = 360 / 255

olddstx = Me.ScaleWidth \ 2
olddsty = Me.ScaleHeight \ 2

For x = 0 To 255 Step 2
dstX = Cos((x * mak) / 180 * PI)
dstY = Sin((x * mak) / 180 * PI)
dstX = (Board.ScaleWidth \ 2) + (dstX * (InData(x) * 0.55))
dstY = (Board.ScaleHeight \ 2) + (dstY * (InData(x) * 0.55))
cr = InData(x) * 1.6
cg = InData(x) * 1.2
cb = InData(x) * 0.8
Board.DrawWidth = mak * 2.5
Board.Line (Board.ScaleWidth \ 2, Board.ScaleHeight \ 2)-(dstX, dstY), RGB(cr, cg, cb)
Board.Line (dstX, dstY)-(olddstx, olddsty), RGB(cr, cg * 0.7, cb * 0.7)
Board.DrawWidth = 1
olddstx = dstX
olddsty = dstY
Next x

Case vLaser

    For x = 0 To 255
         Board.Line (Board.ScaleWidth / 2, 0)-(x * (Board.ScaleWidth / 255), InData(x * 2) * 1.5), RGB(0, InData(x * 2) * 1.7, 0)
    Next x

Case vWalls

    For x = 0 To 255
         Board.Line (Board.ScaleWidth / 2, 0)-(x * (Board.ScaleWidth / 255), InData(x * 2) * 1.5), RGB(0, InData(x * 2) * 1.7, 0), BF
    Next x
    
Case vSLaser

LsrC(LsrI) = 255
LsrX(LsrI) = LsrI * (Board.ScaleWidth / 255)
LsrV(LsrI) = InData(255)

LsrI = LsrI + LsrIStp
If LsrI > 252 Then LsrIStp = -2
If LsrI < 3 Then LsrIStp = 2

    x = LsrI
    Board.Line (Board.ScaleWidth / 2, 0)-(x * (Board.ScaleWidth / 255), InData(255) * 1.5), vbGreen

For x = 0 To 255
Board.Line (Board.ScaleWidth / 2, 0)-(x * (Board.ScaleWidth / 255), LsrV(x) * 1.5), RGB(0, LsrC(x), 0)
LsrC(x) = LsrC(x) - 10
If LsrC(x) < 0 Then LsrC(x) = 0
Next x

Case vShapes

    Dim destX, destY, shpV, destColor
    
    Randomize
    
    For x = 0 To 255 Step 5
    destX = ShpX(x)
    destY = ShpY(x)
    
    shpV = InData(x * 2) * 0.3

    Select Case ShpC(x)
    Case 0
    destColor = RGB(InData(x * 2) * 1.8, 0, 0)
    Case 1
    destColor = RGB(0, InData(x * 2) * 1.8, 0)
    Case 2
    destColor = RGB(0, 0, InData(x * 2) * 1.8)
    End Select
    
    Select Case ShpT(x)
    Case 0 'circle
    Board.Circle (destX, destY), shpV, destColor
    Case 1 'empty box
    Board.Line (destX - shpV / 2, destY - shpV / 2)-(destX + shpV / 2, destY + shpV / 2), destColor, B
    Case 2 'filled box
    Board.Line (destX - shpV / 2, destY - shpV / 2)-(destX + shpV / 2, destY + shpV / 2), destColor, BF
    End Select
    
    If InData(x * 2) > 240 Then
    MakeRandMove 245 - InData(x * 2), x
    End If
    
    Next x
    
Case vGDream: DoDream (0)
Case vYDream: DoDream (1)
Case vADream: DoDream (2)
Case vPDream: DoDream (3)
Case vGrDream: DoDream (4)
Case vBDream: DoDream (5)
Case vRDream: DoDream (6)

Case vRainDream

If DreamCnt > 10 Then
DreamCnt = 0
DreamClr = DreamClr + 1: If DreamClr > 6 Then DreamClr = 0
Else
DreamCnt = DreamCnt + 1
End If

DoDream DreamClr

End Select
        
If cmbV.ListIndex < vGDream And cmbV.ListIndex < vRainDream Then
Board.CurrentX = Board.ScaleWidth \ 2 - Board.TextWidth("FPS: " & TFPS) \ 2
Board.CurrentY = 10
Board.Print "FPS: " & TFPS
End If

FPS = FPS + 1
End Function

Function DoDream(Kind As Integer)
    Dim Stp2 As Long, dY2 As Long
    Dim x As Long
    
    Stp2 = Board.ScaleWidth \ 255

    Dim Max, Min
    Min = 2
    Max = 2

        For x = 0 To MoveM
        intX = (Board.ScaleWidth - 1) * Rnd
        intY = (Board.ScaleHeight - 1) * Rnd
        
        If intX < Board.ScaleWidth / 2 And intY < Board.ScaleHeight / 2 Then intI = -2: intJ = Rand(-Max, -Min)
        If intX > Board.ScaleWidth / 2 And intY > Board.ScaleHeight / 2 Then intI = 2: intJ = Rand(Min, Max)
        
        If intX < Board.ScaleWidth / 2 And intY > Board.ScaleHeight / 2 Then intI = Rand(-Max, -Min): intJ = Rand(Min, Max)
        If intX > Board.ScaleWidth / 2 And intY < Board.ScaleHeight / 2 Then intI = Rand(Min, Max): intJ = Rand(-Max, -Min)
        
        Call BitBlt(Board.hDC, intX + intI, intY + intJ, 88, 88, Board.hDC, intX, intY, vbSrcCopy)
        
        Next x
        
        Dim VU As Double, amp, doub
        
        doub = 90
        amp = 1.4
        
        VU = InData(0) - doub
        VU = ((Board.ScaleHeight / 2) - (VU)) * amp
        Board.CurrentY = VU
        Board.CurrentX = 0
        
        Dim CLR
        
        CLR = InData(Int(Rnd * 511))
        CLR = CLR - 100
        If CLR < 0 Then CLR = 10
        MoveM = CLR * (Board.ScaleWidth / 255)
        CLR = CLR * 3
        CLR = CLR + 50
        
        Select Case Kind
        Case 0 'green dream
        CLR = RGB(0, CLR, 0)
        Case 1 'yellow dream
        CLR = RGB(CLR, CLR, 0)
        Case 2 'aqua dream
        CLR = RGB(0, CLR, CLR)
        Case 3 'purple dream
        CLR = RGB(CLR, 0, CLR)
        Case 4 'grey dream
        CLR = RGB(CLR, CLR, CLR)
        Case 5 'blue dream
        CLR = RGB(0, 0, CLR)
        Case 6 'red dream
        CLR = RGB(CLR, 0, 0)
        End Select
        
    For x = 0 To 255
        dY2 = x * Stp2
        VU = InData(x * 2) - doub
        VU = ((Board.ScaleHeight / 2) - (VU)) * amp
        Board.Line (Board.CurrentX, Board.CurrentY)-(dY2 * 2, VU), CLR
    Next x
End Function

Function OTrim(num As Long)
OTrim = num
If OTrim < 0 Then OTrim = 0
End Function

Private Sub tmrFPS_Timer()
TFPS = FPS
FPS = 0
End Sub
