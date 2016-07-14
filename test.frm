VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9300
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   9300
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1680
      Left            =   3840
      Picture         =   "test.frx":0000
      ScaleHeight     =   1680
      ScaleWidth      =   3360
      TabIndex        =   0
      Top             =   2280
      Width           =   3360
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
                                   
Private Type BITMAP '14 bytes
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type

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
Private Sub Form_Load()
    TransparentPaint Me, Picture1, 0, 0
    
End Sub
