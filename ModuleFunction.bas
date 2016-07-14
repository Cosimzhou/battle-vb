Attribute VB_Name = "ModuleFunction"
Option Explicit
#If Win32 Then
  Private Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long
#Else
  Private Declare Function sndPlaySound Lib "MMSYSTEM" (lpszSoundName As Any, ByVal uFlags%) As Integer
#End If

Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

'*********************************************************************
' Flag values for wFlags parameter.
'*********************************************************************

Public Const SND_SYNC = &H0        ' Play synchronously (default).
Public Const SND_ASYNC = &H1       ' Play asynchronously (see
                                   ' note below).
Public Const SND_NODEFAULT = &H2   ' Do not use default sound.
Public Const SND_MEMORY = &H4      ' lpszSoundName points to a
                                   ' memory file.
Public Const SND_LOOP = &H8        ' Loop the sound until next
                                   ' sndPlaySound.
Public Const SND_NOSTOP = &H10     ' Do not stop any currently
                                   ' playing sound.
Private Const FILESIZEOFAPP2 = 20 * SectionLen
                                   
Public Type BITMAP '14 bytes
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type

Type Bullet
    x As Integer
    y As Integer
    FANG As Integer
    Show As Boolean
    TYPE As Integer
    VD As Integer
End Type

Type Tank
    x As Integer
    y As Integer
'    Level As Integer
    FANG As Integer '方向
    FHP As Integer '命
    FER As Integer '方向速度
    TYPE As Integer
    Show As Boolean '可见
    ProWa As Boolean
    Shield As Boolean
    ShieldID As Integer
    Move As Boolean
End Type

Type Movie
    x As Integer
    y As Integer
    Show As Boolean
    No As Integer
    NUM As Integer
End Type

Type BonusRC
    x As Integer
    y As Integer
    W As Integer
    H As Integer
    T As Integer
    Show As Boolean
End Type

Type Configure
    Player_UP(0 To 1) As Integer
    Player_DOWN(0 To 1) As Integer
    Player_LEFT(0 To 1) As Integer
    Player_RIGHT(0 To 1) As Integer
    Player_SHOOT(0 To 1) As Integer
    DEPTH As Integer
    CLEVER As Integer
End Type

'*********************************************************************
' Plays a wave file from a resource.
'*********************************************************************

Public Sub PlayWaveRes(vntResourceID As Variant, Optional vntFlags)
    '-----------------------------------------------------------------
    ' WARNING:  If you want to play sound files asynchronously in
    '           Win32, then you MUST change bytSound() from a local
    '           variable to a module-level or static variable. Doing
    '           this prevents your array from being destroyed before
    '           sndPlaySound is complete. If you fail to do this, you
    '           will pass an invalid memory pointer, which will cause
    '           a GPF in the Multimedia Control Interface (MCI).
    '-----------------------------------------------------------------
    Dim bytSound() As Byte ' Always store binary data in integer arrays!
    
    bytSound = LoadResData(vntResourceID, "WAVE")
    
    If IsMissing(vntFlags) Then
       vntFlags = SND_SYNC Or SND_NODEFAULT Or SND_MEMORY
    End If
    
    If (vntFlags And SND_MEMORY) = 0 Then
       vntFlags = vntFlags Or SND_MEMORY
    End If
    
    sndPlaySound bytSound(0), vntFlags
End Sub
' *************************************************************************
'   Paints a bitmap on a given surface using the surface backcolor
'   everywhere lngMaskColor appears on the picSource bitmap
' ************************************************************************
Sub TransparentPaint(objDest As Object, picSource As StdPicture, _
        ByVal lngX As Long, ByVal lngY As Long, Optional ByVal lngHeight, _
        Optional ByVal lngWidth, Optional ByVal lngSX = 0, Optional ByVal lngSY = 0, _
        Optional ByVal lngMaskColor As Long = 0)
    '   This   sub   uses   a   bunch   of   variables,   so   let's   declare   and   explain
    '   them   in   advance...
    Dim lngSrcDC As Long             'Source   bitmap
    Dim lngSaveDC As Long           'Copy   of   Source   bitmap
    Dim lngMaskDC     As Long           'Monochrome   Mask   bitmap
    Dim lngInvDC     As Long             'Monochrome   Inverse   of   Mask   bitmap
    Dim lngNewPicDC     As Long       'Combination   of   Source   &   Background   bmps
    Dim bmpSource     As BITMAP       'Description   of   the   Source   bitmap
    Dim hResultBmp     As Long         'Combination   of   source   &   background
    Dim hSaveBmp     As Long             'Copy   of   Source   bitmap
    Dim hMaskBmp     As Long             'Monochrome   Mask   bitmap
    Dim hInvBmp     As Long               'Monochrome   Inverse   of   Mask   bitmap
    Dim hSrcPrevBmp     As Long       'Holds   prev   bitmap   in   source   DC
    Dim hSavePrevBmp     As Long     'Holds   prev   bitmap   in   saved   DC
    Dim hDestPrevBmp     As Long     'Holds   prev   bitmap   in   destination   DC
    Dim hMaskPrevBmp     As Long     'Holds   prev   bitmap   in   the   mask   DC
    Dim hInvPrevBmp     As Long       'Holds   prev   bitmap   in   inverted   mask   DC
    Dim lngOrigScaleMode&           'Holds   the   original   ScaleMode
    Dim lngOrigColor&                   'Holds   original   backcolor   from   source   DC
    '   Set   ScaleMode   to   pixels   for   Windows   GDI
    lngOrigScaleMode = objDest.ScaleMode
    objDest.ScaleMode = vbPixels
    '   Load   the   source   bitmap   to   get   its   width   (bmpSource.bmWidth)
    '   and   height   (bmpSource.bmHeight)
    GetObject picSource, Len(bmpSource), bmpSource
    If Not IsMissing(lngHeight) Then
        bmpSource.bmHeight = lngHeight
        bmpSource.bmWidth = lngWidth
    End If
    '   Create   compatible   device   contexts   (DC's)   to   hold   the   temporary
    '   bitmaps   used   by   this   sub
    lngSrcDC = CreateCompatibleDC(objDest.hdc)
    lngSaveDC = CreateCompatibleDC(objDest.hdc)
    lngMaskDC = CreateCompatibleDC(objDest.hdc)
    lngInvDC = CreateCompatibleDC(objDest.hdc)
    lngNewPicDC = CreateCompatibleDC(objDest.hdc)
    '   Create   monochrome   bitmaps   for   the   mask-related   bitmaps
    hMaskBmp = CreateBitmap(bmpSource.bmWidth, bmpSource.bmHeight, 1, 1, ByVal 0&)
    hInvBmp = CreateBitmap(bmpSource.bmWidth, bmpSource.bmHeight, 1, 1, ByVal 0&)
    '   Create   color   bitmaps   for   the   final   result   and   the   backup   copy
    '   of   the   source   bitmap
    hResultBmp = CreateCompatibleBitmap(objDest.hdc, bmpSource.bmWidth, bmpSource.bmHeight)
    hSaveBmp = CreateCompatibleBitmap(objDest.hdc, bmpSource.bmWidth, bmpSource.bmHeight)
    '   Select   bitmap   into   the   device   context   (DC)
    hSrcPrevBmp = SelectObject(lngSrcDC, picSource)
    hSavePrevBmp = SelectObject(lngSaveDC, hSaveBmp)
    hMaskPrevBmp = SelectObject(lngMaskDC, hMaskBmp)
    hInvPrevBmp = SelectObject(lngInvDC, hInvBmp)
    hDestPrevBmp = SelectObject(lngNewPicDC, hResultBmp)
    '   Make   a   backup   of   source   bitmap   to   restore   later
    BitBlt lngSaveDC, 0, 0, bmpSource.bmWidth, bmpSource.bmHeight, lngSrcDC, lngSX, lngSY, vbSrcCopy
    
    '   Create   the   mask   by   setting   the   background   color   of   source   to
    '   transparent   color,   then   BitBlt'ing   that   bitmap   into   the   mask
    '   device   context
    lngOrigColor = SetBkColor(lngSrcDC, lngMaskColor)
    BitBlt lngMaskDC, 0, 0, bmpSource.bmWidth, bmpSource.bmHeight, lngSrcDC, lngSX, lngSY, vbSrcCopy
    '   Restore   the   original   backcolor   in   the   device   context
    SetBkColor lngSrcDC, lngOrigColor
    '   Create   an   inverse   of   the   mask   to   AND   with   the   source   and   combine
    '   it   with   the   background
    BitBlt lngInvDC, 0, 0, bmpSource.bmWidth, bmpSource.bmHeight, lngMaskDC, 0, 0, vbNotSrcCopy
    '   Copy   the   background   bitmap   to   the   new   picture   device   context
    '   to   begin   creating   the   final   transparent   bitmap
    BitBlt lngNewPicDC, 0, 0, bmpSource.bmWidth, bmpSource.bmHeight, objDest.hdc, lngX, lngY, vbSrcCopy
'    BitBlt
    '   AND   the   mask   bitmap   with   the   result   device   context   to   create
    '   a   cookie   cutter   effect   in   the   background   by   painting   the   black
    '   area   for   the   non-transparent   portion   of   the   source   bitmap
    BitBlt lngNewPicDC, 0, 0, bmpSource.bmWidth, bmpSource.bmHeight, lngMaskDC, 0, 0, vbSrcAnd
    '   AND   the   inverse   mask   with   the   source   bitmap   to   turn   off   the   bits
    '   associated   with   transparent   area   of   source   bitmap   by   making   it
    '   black
    BitBlt lngSrcDC, lngSX, lngSY, bmpSource.bmWidth, bmpSource.bmHeight, lngInvDC, 0, 0, vbSrcAnd
    '   XOR   the   result   with   the   source   bitmap   to   replace   the   mask   color
    '   with   the   background   color
    BitBlt lngNewPicDC, 0, 0, bmpSource.bmWidth, bmpSource.bmHeight, lngSrcDC, lngSX, lngSY, vbSrcPaint
    '   Paint   the   transparent   bitmap   on   source   surface
    BitBlt objDest.hdc, lngX, lngY, bmpSource.bmWidth, bmpSource.bmHeight, lngNewPicDC, 0, 0, vbSrcCopy
    '   Restore   backup   of   bitmap
    BitBlt lngSrcDC, lngSX, lngSY, bmpSource.bmWidth, bmpSource.bmHeight, lngSaveDC, 0, 0, vbSrcCopy
    '   Restore   the   original   objects   by   selecting   their   original   values
    SelectObject lngSrcDC, hSrcPrevBmp
    SelectObject lngSaveDC, hSavePrevBmp
    SelectObject lngNewPicDC, hDestPrevBmp
    SelectObject lngMaskDC, hMaskPrevBmp
    SelectObject lngInvDC, hInvPrevBmp
    '   Free   system   resources   created   by   this   sub
    DeleteObject hSaveBmp
    DeleteObject hMaskBmp
    DeleteObject hInvBmp
    DeleteObject hResultBmp
    DeleteDC lngSrcDC
    DeleteDC lngSaveDC
    DeleteDC lngInvDC
    DeleteDC lngMaskDC
    DeleteDC lngNewPicDC
    '   Restores   the   ScaleMode   to   its   original   value
    objDest.ScaleMode = lngOrigScaleMode
End Sub
'‘Sub slidetext("任务失败！", 150, 200, 12, vbBlue, Picture1)

'End Sub

Public Sub ReleaseMap()
    Dim APP2() As Byte
    Dim Counter As Long
    APP2 = LoadResData("MAPFILE", "MAPBIN")

    If Dir("Map.bin") <> "" Then Exit Sub

    Open "Map.bin" For Binary Access Write As #1
        For Counter = 0 To FILESIZEOFAPP2 - 1
            Put #1, , APP2(Counter)
        Next Counter
    Close #1
    
    APP2 = LoadResData("GNINFO", "TEXT")
    
    If Dir("gninfo.txt") <> "" Then Exit Sub

    Open "gninfo.txt" For Binary Access Write As #1
        For Counter = 0 To FILESIZEOFAPP2 - 1
            Put #1, , APP2(Counter)
        Next Counter
    Close #1

End Sub

Sub Main()
    On Error Resume Next
    Dim Setting As Configure
    Dim APP2() As Byte
    Dim App_Path As String
    Dim Counter As Long
    'App_Path = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\")
    MainForm.App_Path = App_Path
    If Dir$(App_Path & "Tank.cfg") = "" Then
        MsgBox "文件损坏,系统将自动重新安装."
        With Setting
            .Player_UP(0) = vbKeyUp
            .Player_DOWN(0) = vbKeyDown
            .Player_LEFT(0) = vbKeyLeft
            .Player_RIGHT(0) = vbKeyRight
            .Player_SHOOT(0) = vbKeySpace
            .Player_UP(1) = vbKeyW
            .Player_DOWN(1) = vbKeyS
            .Player_LEFT(1) = vbKeyA
            .Player_RIGHT(1) = vbKeyD
            .Player_SHOOT(1) = vbKeyJ
            .DEPTH = 3
            .CLEVER = 3
        End With
        Open App_Path & "Tank.cfg" For Binary Access Write As 1 Len = Len(Setting)
        Put #1, , Setting
        Close
        ReleaseMap
        
    '    BANG                  WAVE    DISCARDABLE     "E:\sound\Bang.wav"
    '     FARE                  WAVE    DISCARDABLE     "E:\sound\Fanfare.wav"
    '     FIRE                  WAVE    DISCARDABLE     "E:\sound\Gunfire.wav"
    '     HIT                   WAVE    DISCARDABLE     "E:\sound\hit.wav"
    '     PEOW                  WAVE    DISCARDABLE     "E:\sound\Peow.wav"
    '     MAPFILE               MAPBIN  DISCARDABLE     "E:\sound\Map.bin"
        APP2 = LoadResData("BANG", "WAVE")
        Counter = UBound(APP2)
        ReDim Preserve APP2(Counter) As Byte
        Open App_Path & "bang.wav " For Binary Access Write As #1
            Put #1, , APP2
        Close #1
        
        APP2 = LoadResData("FARE", "WAVE")
        Counter = UBound(APP2)
        ReDim Preserve APP2(Counter) As Byte
        Open App_Path & "Fanfare.wav " For Binary Access Write As #1
            Put #1, , APP2
        Close #1
        
        APP2 = LoadResData("FIRE", "WAVE")
        Counter = UBound(APP2)
        ReDim Preserve APP2(Counter) As Byte
        Open App_Path & "fire.wav " For Binary Access Write As #1
            Put #1, , APP2
        Close #1
        
        APP2 = LoadResData("hit", "WAVE")
        Counter = UBound(APP2)
        ReDim Preserve APP2(Counter) As Byte
        Open App_Path & "hit.wav " For Binary Access Write As #1
            Put #1, , APP2
        Close #1
        
        APP2 = LoadResData("peow", "WAVE")
        Counter = UBound(APP2)
        ReDim Preserve APP2(Counter) As Byte
        Open App_Path & "peow.wav " For Binary Access Write As #1
            Put #1, , APP2
        Close #1
        
        
        MkDir App_Path & "Image"
        MkDir App_Path & "Temp"

    End If
    Open App_Path & "Tank.cfg" For Binary Access Read As 1
    Get #1, , Setting
    Close #1
    MainForm.Show
End Sub
