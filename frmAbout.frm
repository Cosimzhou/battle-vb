VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "关于坦克大战"
   ClientHeight    =   3690
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5865
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":08CA
   ScaleHeight     =   246
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   391
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   5280
      Top             =   0
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      Height          =   495
      Left            =   4440
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label CmdOK 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "确定"
      BeginProperty Font 
         Name            =   "华文行楷"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   495
      Left            =   4440
      TabIndex        =   6
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   120
      Picture         =   "frmAbout.frx":59F3
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   720
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1095
      Left            =   240
      Picture         =   "frmAbout.frx":62BD
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "cosimzhou@hotmail.com"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9
         Charset         =   0
         Weight          =   300
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   1680
      MousePointer    =   1  'Arrow
      TabIndex        =   5
      Top             =   3285
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "E_mail:"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   10.5
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   3240
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   6
      X2              =   376.933
      Y1              =   163
      Y2              =   163
   End
   Begin VB.Label lblD 
      BackStyle       =   0  'Transparent
      Caption         =   "本游戏版权归Sestall Studio所有，任何人不得擅自转让、出售本游戏！"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   1170
      Left            =   1560
      TabIndex        =   0
      Top             =   1080
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "坦克大战"
      BeginProperty Font 
         Name            =   "华文彩云"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   585
      Left            =   2040
      TabIndex        =   2
      Top             =   120
      Width           =   2805
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   7
      X2              =   376.933
      Y1              =   164
      Y2              =   164
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "1.03.02版"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   225
      Left            =   2640
      TabIndex        =   3
      Top             =   780
      Width           =   1365
   End
   Begin VB.Label lblDi 
      BackStyle       =   0  'Transparent
      Caption         =   "该游戏如有问题，请与供应商联系。"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   345
      Left            =   360
      TabIndex        =   1
      Top             =   2625
      Width           =   3630
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Dim a As Long
Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub CmdOK_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Shape1.BorderColor = vbWhite
    CmdOK.ForeColor = vbRed
End Sub

Private Sub Form_Load()
    Dim Rtn As Long
    Randomize
    Rtn = SetWindowPos(frmAbout.hwnd, -1, 0, 0, 0, 0, 3)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Shape1.BorderColor = vbRed
    CmdOK.ForeColor = &HC0C0&
    Label4.ForeColor = &HFF8080
End Sub

Private Sub Label4_Click()
    ShellExecute 0&, vbNullString, "mailto:ufis8@hotmail.com", vbNullString, vbNullString, 0
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Label4.ForeColor = vbBlue
End Sub

Private Sub Timer1_Timer()
    a = Int(Rnd * 15 + 1)
    If a <> 7 Then
        lblTitle.ForeColor = QBColor(a)
    End If
End Sub
