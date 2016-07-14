Attribute VB_Name = "ModuleMain"
Global Const SRCCOPY = &HCC0020
Global Const SectionLen = 676
Global Const ZHUANQIANG = 1
Global Const SHULIN = 2
Global Const TIEQIANG = 3
Global Const HESHUI = 4
Global Const BINGDI = 5
Const MASKMODE = &H440328

Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Enum Map
    Blank = 0
    Brick = 1
    Tree = 2
    Iron = 3
    Sea = 4
End Enum

Type POINTAPI
    x As Integer
    y As Integer
End Type

Sub Tomin(ByVal x As Long, ByVal y As Long, ByVal maskX As Long, ByVal maskY As Long, ByVal aimX As Long, ByVal aimY As Long, ByVal scHeight As Long, ByVal scWidth As Long, ByVal aimhDC As Long, ByVal sourcehDC As Long)
    BitBlt aimhDC, x, y, scWidth, scHeight, sourcehDC, maskX, maskY, &HD9A02 'MASKMODE
    BitBlt aimhDC, x, y, scWidth, scHeight, sourcehDC, aimX, aimY, &HEE0086 'SRCCOPY
End Sub
