VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "设置"
   ClientHeight    =   3015
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5925
   Icon            =   "frmSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   "敌方坦克"
      Height          =   2295
      Left            =   3720
      TabIndex        =   24
      Top             =   120
      Width           =   2055
      Begin ComctlLib.Slider SldClever 
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1080
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   327682
         LargeChange     =   1
         Max             =   3
         SelStart        =   3
         Value           =   3
      End
      Begin ComctlLib.Slider SldDepth 
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   327682
         LargeChange     =   1
         Max             =   3
         SelStart        =   3
         Value           =   3
      End
      Begin VB.Label lblClever 
         Caption         =   "狡猾程度:狐狸"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label lblDepth 
         Caption         =   "凶悍程度:饕餮"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "玩家二"
      Height          =   2295
      Left            =   1920
      TabIndex        =   13
      Top             =   120
      Width           =   1695
      Begin VB.TextBox txtKeyVal 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   9
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox txtKeyVal 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   8
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtKeyVal 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   7
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtKeyVal 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   6
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtKeyVal 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   5
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "开火"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   23
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "向右"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   21
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "向左"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "向下"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "向上"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "玩家一"
      Height          =   2295
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1695
      Begin VB.TextBox txtKeyVal 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   4
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox txtKeyVal 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   3
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtKeyVal 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   2
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtKeyVal 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   1
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtKeyVal 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   0
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "开火"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "向右"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "向左"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "向下"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "向上"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "取消"
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "确定"
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   2520
      Width           =   1215
   End
End
Attribute VB_Name = "frmSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const DEPTHSTR As String = "饭桐一般精锐饕餮"
Const CLEVERSTR As String = "爬虫跳蚤猎犬狐狸"

Dim Setting As Configure

Private Sub Slider1_Click()

End Sub

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Open MainForm.App_Path & "Tank.cfg" For Binary Access Read As 1 Len = Len(Setting)
        Get #1, , Setting
    Close #1
    txtKeyVal(0).Text = Setting.Player_UP(0)
    txtKeyVal(1).Text = Setting.Player_DOWN(0)
    txtKeyVal(2).Text = Setting.Player_LEFT(0)
    txtKeyVal(3).Text = Setting.Player_RIGHT(0)
    txtKeyVal(4).Text = Setting.Player_SHOOT(0)
    txtKeyVal(5).Text = Setting.Player_UP(1)
    txtKeyVal(6).Text = Setting.Player_DOWN(1)
    txtKeyVal(7).Text = Setting.Player_LEFT(1)
    txtKeyVal(8).Text = Setting.Player_RIGHT(1)
    txtKeyVal(9).Text = Setting.Player_SHOOT(1)
    SldDepth.Value = Setting.DEPTH
    SldClever.Value = Setting.CLEVER
End Sub

Private Sub OKButton_Click()
    Setting.Player_UP(0) = Val(txtKeyVal(0).Text)
    Setting.Player_DOWN(0) = Val(txtKeyVal(1).Text)
    Setting.Player_LEFT(0) = Val(txtKeyVal(2).Text)
    Setting.Player_RIGHT(0) = Val(txtKeyVal(3).Text)
    Setting.Player_SHOOT(0) = Val(txtKeyVal(4).Text)
    Setting.Player_UP(1) = Val(txtKeyVal(5).Text)
    Setting.Player_DOWN(1) = Val(txtKeyVal(6).Text)
    Setting.Player_LEFT(1) = Val(txtKeyVal(7).Text)
    Setting.Player_RIGHT(1) = Val(txtKeyVal(8).Text)
    Setting.Player_SHOOT(1) = Val(txtKeyVal(9).Text)
    Setting.DEPTH = SldDepth.Value
    Setting.CLEVER = SldClever.Value
    Open MainForm.App_Path & "Tank.cfg" For Binary Access Write As 1 Len = Len(Setting)
        Put #1, , Setting
    Close #1
    MainForm.RefreshSetting
    Unload Me
End Sub

Private Sub SldClever_Click()
    lblClever.Caption = "狡猾程度:" & Mid(CLEVERSTR, SldClever.Value * 2 + 1, 2)
End Sub

Private Sub SldDepth_Click()
    lblDepth.Caption = "凶悍程度:" & Mid(DEPTHSTR, SldDepth.Value * 2 + 1, 2)
End Sub

Private Sub txtKeyVal_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    txtKeyVal(Index).Text = KeyCode
End Sub
