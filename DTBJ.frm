VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form DTBJ 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "地图编辑器"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9090
   Icon            =   "DTBJ.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   437
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   606
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   255
      Left            =   6600
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   9
      Top             =   3720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox PicDX 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   5
      Left            =   8640
      Picture         =   "DTBJ.frx":030A
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   8
      Top             =   6240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox PicDX 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   0
      Left            =   6600
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   7
      Top             =   6240
      Visible         =   0   'False
      Width           =   300
   End
   Begin MSComDlg.CommonDialog CoDi 
      Left            =   6120
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "*.bmp"
      Filter          =   "位图文件(.bmp)|*.bmp"
      InitDir         =   "App.Path"
   End
   Begin VB.PictureBox PicIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   1
      Left            =   6600
      Picture         =   "DTBJ.frx":064E
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   6
      Top             =   6120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.ListBox List 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2670
      ItemData        =   "DTBJ.frx":1290
      Left            =   6600
      List            =   "DTBJ.frx":1292
      TabIndex        =   5
      Top             =   120
      Width           =   2295
   End
   Begin VB.PictureBox PicDX 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   2
      Left            =   7320
      Picture         =   "DTBJ.frx":1294
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   4
      Top             =   6240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox PicDX 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   4
      Left            =   8280
      Picture         =   "DTBJ.frx":18D6
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   3
      Top             =   6240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox PicDX 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   1
      Left            =   6960
      Picture         =   "DTBJ.frx":1C18
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   2
      Top             =   6240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox PicDX 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   3
      Left            =   7920
      Picture         =   "DTBJ.frx":1F5A
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   1
      Top             =   6240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox PicDT 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      DrawWidth       =   2
      Height          =   6300
      Left            =   120
      MousePointer    =   2  'Cross
      ScaleHeight     =   416
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   416
      TabIndex        =   0
      Top             =   120
      Width           =   6300
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         FillColor       =   &H00FF0000&
         Height          =   240
         Left            =   -240
         Shape           =   1  'Square
         Top             =   120
         Width           =   240
      End
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2175
      Left            =   6600
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Menu Game 
      Caption         =   "游戏(&G)"
      Begin VB.Menu Option 
         Caption         =   "选择关卡(&O)"
         Shortcut        =   {F2}
      End
      Begin VB.Menu Save 
         Caption         =   "储存关卡(&S)"
         Shortcut        =   {F3}
      End
      Begin VB.Menu cde 
         Caption         =   "-"
      End
      Begin VB.Menu SaveAs 
         Caption         =   "另存为图片(&A)"
      End
      Begin VB.Menu ABCDEFGH 
         Caption         =   "-"
      End
      Begin VB.Menu Clear 
         Caption         =   "恢复关卡原状(&C)"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "帮助(&H)"
      Begin VB.Menu About 
         Caption         =   "关于..."
      End
   End
End
Attribute VB_Name = "DTBJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const SectionLen = 676
Dim DT As Long
Dim x As Long, y As Long
Dim XO As Long, YO As Long
Dim Map(25, 25) As Byte
Dim H As Long
Dim n As Integer
Dim S As String, FS As String
Dim Length As Long
Dim Modify As Boolean

Private Sub About_Click()
    frmAbout.Show modal
End Sub
Private Sub Clear_Click()
    PicDT.Cls
End Sub
Private Sub Form_Load()
    Picture1.Height = PicDX(0).Height
    Picture1.Width = PicDX(0).Width
    List.AddItem "砖墙"
    List.AddItem "铁墙"
    List.AddItem "树林"
    List.AddItem "海洋"
    List.AddItem "冰地"
    List.AddItem "空地"
    FS = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\")
    n = 1
    Read
    Begin
    DTBJ.Caption = "地图编辑器—" + "任务" & n & "----" & S
End Sub
Sub Begin()
    PicDT.Cls
    Tomin 192, 384, 0, 32, 0, 0, 32, 32, PicDT.hDC, PicIcon(1).hDC
    For i = 0 To 25
        For j = 0 To 25
            If Map(i, j) > 0 And Map(i, j) < 6 Then
                If Map(i, j) = 2 Then
                    Tomin i * 16, j * 16, 16, 0, 0, 0, 16, 16, PicDT.hDC, PicDX(2).hDC
                Else
                    BitBlt PicDT.hDC, i * 16, j * 16, 16, 16, PicDX(Map(i, j)).hDC, 0, 0, SRCCOPY
                End If
            End If
        Next j
    Next i
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim Temp As VbMsgBoxResult
    If Modify Then
        Temp = MsgBox("是否保存第" & n & "关的地图？", vbYesNoCancel + vbQuestion, "存盘提示")
        Select Case Temp
        Case vbYes
           Call Save_Click
        Case vbCancel
           Cancel = True
        End Select
    End If
End Sub


Private Sub List_Click()
    Select Case List.Text
    Case "砖墙"
        DT = 1
        BitBlt Picture1.hDC, 0, 0, 16, 16, PicDX(1).hDC, 0, 0, SRCCOPY
        Image1.Picture = Picture1.Image
    Case "铁墙"
        DT = 3
        BitBlt Picture1.hDC, 0, 0, 16, 16, PicDX(3).hDC, 0, 0, SRCCOPY
        Image1.Picture = Picture1.Image
    Case "树林"
        DT = 2
        BitBlt Picture1.hDC, 0, 0, 16, 16, PicDX(2).hDC, 0, 0, SRCCOPY
        Image1.Picture = Picture1.Image
    Case "海洋"
        DT = 4
        BitBlt Picture1.hDC, 0, 0, 16, 16, PicDX(4).hDC, 0, 0, SRCCOPY
        Image1.Picture = Picture1.Image
    Case "冰地"
        DT = 5
        BitBlt Picture1.hDC, 0, 0, 16, 16, PicDX(5).hDC, 0, 0, SRCCOPY
        Image1.Picture = Picture1.Image
    Case "空地"
        DT = 0
        BitBlt Picture1.hDC, 0, 0, 16, 16, PicDX(0).hDC, 0, 0, SRCCOPY
        Image1.Picture = Picture1.Image
    End Select
End Sub
Private Sub Option_Click()
    On Error Resume Next
    t% = FileLen(FS & "Map.bin") \ SectionLen
    n = Val(InputBox("请输入选择关卡的序列号(1-" & t & ",共" & t & "关):" & vbLf & "    注:也可输入" & t + 1 & "以新建一关.", "选择关卡", IIf((n + 1) Mod t = 0, t, (n + 1) Mod t)))
    If n <= 0 Then Exit Sub
    If n > t + 1 Then n = t + 1
    Read
    If n > t Then
        Map(11, 25) = ZHUANQIANG
        Map(11, 24) = ZHUANQIANG
        Map(11, 23) = ZHUANQIANG
        Map(12, 23) = ZHUANQIANG
        Map(13, 23) = ZHUANQIANG
        Map(14, 24) = ZHUANQIANG
        Map(14, 25) = ZHUANQIANG
        Map(14, 23) = ZHUANQIANG
    End If
    PicDT.Cls
    DTBJ.Caption = "地图编辑器—" + "任务" & n
    Begin
End Sub
Private Sub PicDT_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Click
    End If
End Sub
Private Sub PicDT_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    XO = x \ 16
    YO = y \ 16
    PicDT.ToolTipText = "(" & XO & "," & YO & ")"
    Shape1.Top = YO * 16
    Shape1.Left = XO * 16
    
    If Button = 1 Then
        Click
    End If
End Sub
Sub Click()
    Dim i As Integer
    If XO < 0 Or XO > 25 Or YO < 0 Or YO > 25 Then Exit Sub
    If DT = 2 Then
        BitBlt PicDT.hDC, XO * 16, YO * 16, 16, 16, PicDX(0).hDC, 0, 0, SRCCOPY
        Tomin XO * 16, YO * 16, 16, 0, 0, 0, 16, 16, PicDT.hDC, PicDX(2).hDC
        Map(XO, YO) = 2
    Else
        BitBlt PicDT.hDC, XO * 16, YO * 16, 16, 16, PicDX(DT).hDC, 0, 0, SRCCOPY
        Map(XO, YO) = DT
    End If
    Modify = True
End Sub

Private Sub Save_Click()
    Dim strTmp As String, Strmp As String, i As Integer
    If MsgBox("是否修改关名?", vbInformation + vbYesNo) = vbYes Then
        strTmp = InputBox("请输入新的关名:", "更改关名", S)
        If strTmp <> "" Then S = strTmp
        Strmp = ""
        Open FS & "gninfo.txt" For Input As #1
            i = 1
            Do Until EOF(1)
                Line Input #1, strTmp
                If i = n Then
                    Strmp = Strmp + S + vbNewLine
                Else
                    Strmp = Strmp + strTmp + vbNewLine
                End If
                i = i + 1
            Loop
            Do While i <= n
                If i = n Then
                    Strmp = Strmp + S + vbNewLine
                Else
                    Strmp = Strmp + vbNewLine
                End If
                i = i + 1
            Loop
        Close
        Open FS & "gninfo.txt" For Output As #1
            Print #1, Strmp
        Close
    End If
    
    For i = 0 To 25
        For j = 0 To 25
            If Map(i, j) <= 0 Or Map(i, j) > 5 Then Map(i, j) = 0
        Next j
    Next i
    Open FS & "Map.bin" For Binary Access Write As #1 Len = SectionLen
        Put #1, (n - 1) * SectionLen + 1, Map
    Close
    Modify = False
End Sub
Sub Read()
    Dim i As Integer
    Open FS & "Map.bin" For Binary Access Read As #1 Len = SectionLen '676
        Get #1, (n - 1) * SectionLen + 1, Map
    Close
    i = 0
    Open FS & "gninfo.txt" For Input As 1
        Do Until EOF(1) Or i = n
            Line Input #1, S
            i = i + 1
        Loop
        If i <> n Then S = "未命名关"
    Close
End Sub
Private Sub SaveAs_Click()
    Dim stmp As String
    CoDi.ShowSave
    CoDi.FileName = LCase(CoDi.FileName)
    If Right(CoDi.FileName, 4) <> ".bmp" Then CoDi.FileName = CoDi.FileName + ".bmp"
    stmp = Dir$(CoDi.FileName)
    If Len(stmp) = 0 Then
        SavePicture PicDT.Image, CoDi.FileName
    ElseIf stmp = Right(CoDi.FileName, Len(stmp)) Then
        If MsgBox("此文件已存在,是否真的要覆盖此文件?", vbYesNo + vbExclamation, "另存为") = vbYes Then SavePicture PicDT.Image, CoDi.FileName
    End If
End Sub
