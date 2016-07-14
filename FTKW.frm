VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form MainForm 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "坦克大战"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7725
   Icon            =   "FTKW.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   433
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   515
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox PicDX 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   5
      Left            =   5040
      Picture         =   "FTKW.frx":1CFA
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   35
      Top             =   4080
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox PicShield 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   3840
      Picture         =   "FTKW.frx":203C
      ScaleHeight     =   480
      ScaleWidth      =   960
      TabIndex        =   33
      Top             =   3840
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox PicWin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4035
      Index           =   1
      Left            =   11640
      Picture         =   "FTKW.frx":2C80
      ScaleHeight     =   4035
      ScaleWidth      =   2580
      TabIndex        =   32
      Top             =   600
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.PictureBox PicWin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   6300
      Index           =   0
      Left            =   8520
      Picture         =   "FTKW.frx":1963C
      ScaleHeight     =   6300
      ScaleWidth      =   6300
      TabIndex        =   31
      Top             =   0
      Visible         =   0   'False
      Width           =   6300
   End
   Begin VB.Timer BornKeeper 
      Interval        =   100
      Left            =   7920
      Top             =   3480
   End
   Begin VB.Timer SleepBonus 
      Enabled         =   0   'False
      Interval        =   7000
      Left            =   7920
      Top             =   3000
   End
   Begin VB.Timer Bullet 
      Interval        =   40
      Left            =   7920
      Top             =   2520
   End
   Begin VB.Timer Update 
      Interval        =   100
      Left            =   7920
      Top             =   1560
   End
   Begin VB.Timer RndEnemy 
      Interval        =   60
      Left            =   7920
      Top             =   2040
   End
   Begin VB.Timer VideoTimer 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   7920
      Top             =   600
   End
   Begin VB.Timer Sleeper 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7920
      Top             =   1080
   End
   Begin VB.Timer RunBonus 
      Interval        =   10000
      Left            =   7920
      Top             =   120
   End
   Begin VB.PictureBox PicBack 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   6300
      Left            =   120
      ScaleHeight     =   416
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   416
      TabIndex        =   9
      Top             =   6360
      Width           =   6300
      Begin VB.Timer Timer8 
         Enabled         =   0   'False
         Interval        =   20000
         Left            =   0
         Top             =   3360
      End
   End
   Begin VB.PictureBox PicMytank 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1680
      Index           =   1
      Left            =   1920
      Picture         =   "FTKW.frx":9A9AE
      ScaleHeight     =   112
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   280
      TabIndex        =   18
      Top             =   2160
      Visible         =   0   'False
      Width           =   4200
   End
   Begin VB.PictureBox PicMytank 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1680
      Index           =   0
      Left            =   1920
      Picture         =   "FTKW.frx":B1972
      ScaleHeight     =   112
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   280
      TabIndex        =   17
      Top             =   480
      Visible         =   0   'False
      Width           =   4200
   End
   Begin VB.PictureBox PicBonus 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   5
      Left            =   4320
      Picture         =   "FTKW.frx":B9836
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox PicBZ 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2400
      Index           =   2
      Left            =   6000
      Picture         =   "FTKW.frx":BA078
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   15
      Top             =   480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox PicBonus 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   4
      Left            =   3840
      Picture         =   "FTKW.frx":BB8BA
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox PicBonus 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   2
      Left            =   2880
      Picture         =   "FTKW.frx":BC0FC
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox PicBonus 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   3
      Left            =   3360
      Picture         =   "FTKW.frx":BC93E
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox PicBonus 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   1
      Left            =   2400
      Picture         =   "FTKW.frx":BD180
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox PicBonus 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   0
      Left            =   1920
      Picture         =   "FTKW.frx":BD9C2
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox Picetank 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3360
      Left            =   1920
      Picture         =   "FTKW.frx":BE204
      ScaleHeight     =   224
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   280
      TabIndex        =   8
      Top             =   4320
      Visible         =   0   'False
      Width           =   4200
   End
   Begin VB.PictureBox PicBZ 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   120
      Index           =   1
      Left            =   4800
      Picture         =   "FTKW.frx":EC148
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   6
      Top             =   3840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox PicIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   1
      Left            =   6000
      Picture         =   "FTKW.frx":EC68A
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   5
      Top             =   2880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox PicDX 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   2
      Left            =   5520
      Picture         =   "FTKW.frx":ED2CC
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   4
      Top             =   3840
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox PicDX 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   0
      Left            =   5280
      Picture         =   "FTKW.frx":ED60E
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   3
      Top             =   3840
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox PicDX 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   3
      Left            =   5280
      Picture         =   "FTKW.frx":ED950
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   2
      Top             =   4080
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox PicIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   0
      Left            =   6000
      Picture         =   "FTKW.frx":EDC92
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   3360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox PicBZ 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   5700
      Index           =   3
      Left            =   0
      Picture         =   "FTKW.frx":EE8D4
      ScaleHeight     =   380
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.PictureBox PicStar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   1920
      Picture         =   "FTKW.frx":112316
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   19
      Top             =   3840
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   6495
      Left            =   6525
      TabIndex        =   22
      Top             =   0
      Width           =   1215
      Begin VB.Label lblLife 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   15.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   600
         TabIndex        =   25
         Top             =   4560
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   1
         Left            =   240
         Picture         =   "FTKW.frx":113758
         Top             =   4560
         Width           =   240
      End
      Begin VB.Label lblLevel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   15.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   24
         Top             =   5640
         Width           =   615
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   240
         Picture         =   "FTKW.frx":113AC8
         Top             =   5400
         Width           =   480
      End
      Begin VB.Label lblLife 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   15.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   600
         TabIndex        =   23
         Top             =   3960
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   0
         Left            =   240
         Picture         =   "FTKW.frx":113E80
         Top             =   3960
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   0
         Left            =   240
         Picture         =   "FTKW.frx":1141F0
         Top             =   240
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   1
         Left            =   720
         Picture         =   "FTKW.frx":114566
         Top             =   240
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   2
         Left            =   240
         Picture         =   "FTKW.frx":1148DC
         Top             =   600
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   3
         Left            =   720
         Picture         =   "FTKW.frx":114C52
         Top             =   600
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   4
         Left            =   240
         Picture         =   "FTKW.frx":114FC8
         Top             =   960
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   5
         Left            =   720
         Picture         =   "FTKW.frx":11533E
         Top             =   960
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   6
         Left            =   240
         Picture         =   "FTKW.frx":1156B4
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   7
         Left            =   720
         Picture         =   "FTKW.frx":115A2A
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   8
         Left            =   240
         Picture         =   "FTKW.frx":115DA0
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   9
         Left            =   720
         Picture         =   "FTKW.frx":116116
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   10
         Left            =   240
         Picture         =   "FTKW.frx":11648C
         Top             =   2040
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   11
         Left            =   720
         Picture         =   "FTKW.frx":116802
         Top             =   2040
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   12
         Left            =   240
         Picture         =   "FTKW.frx":116B78
         Top             =   2400
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   13
         Left            =   720
         Picture         =   "FTKW.frx":116EEE
         Top             =   2400
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   14
         Left            =   240
         Picture         =   "FTKW.frx":117264
         Top             =   2760
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   15
         Left            =   720
         Picture         =   "FTKW.frx":1175DA
         Top             =   2760
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   16
         Left            =   240
         Picture         =   "FTKW.frx":117950
         Top             =   3120
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   17
         Left            =   720
         Picture         =   "FTKW.frx":117CC6
         Top             =   3120
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   18
         Left            =   240
         Picture         =   "FTKW.frx":11803C
         Top             =   3480
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   19
         Left            =   720
         Picture         =   "FTKW.frx":1183B2
         Top             =   3480
         Width           =   240
      End
   End
   Begin MSComDlg.CommonDialog CoDi 
      Left            =   6240
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "位图文件(.bmp)|*.bmp"
   End
   Begin VB.PictureBox PicBonus 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   6
      Left            =   4800
      Picture         =   "FTKW.frx":118728
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox PicBonus 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   7
      Left            =   5280
      Picture         =   "FTKW.frx":11936C
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox PicBonus 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   8
      Left            =   5760
      Picture         =   "FTKW.frx":119BAE
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox PicDX 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   4
      Left            =   5520
      Picture         =   "FTKW.frx":11A3F0
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   27
      Top             =   4080
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox PicSplash 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3330
      Left            =   8520
      Picture         =   "FTKW.frx":11A732
      ScaleHeight     =   3330
      ScaleWidth      =   5640
      TabIndex        =   28
      Top             =   0
      Visible         =   0   'False
      Width           =   5640
   End
   Begin VB.PictureBox PicGameOver 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   8520
      Picture         =   "FTKW.frx":124ABC
      ScaleHeight     =   2415
      ScaleWidth      =   3750
      TabIndex        =   29
      Top             =   0
      Visible         =   0   'False
      Width           =   3750
   End
   Begin VB.PictureBox PicBZ 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   120
      Index           =   0
      Left            =   4800
      Picture         =   "FTKW.frx":12ED7A
      ScaleHeight     =   120
      ScaleWidth      =   480
      TabIndex        =   30
      Top             =   3960
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox PicDX 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   1
      Left            =   4800
      Picture         =   "FTKW.frx":12F2BC
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   1
      Top             =   4080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox PicNum 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   0
      Picture         =   "FTKW.frx":12F8FE
      ScaleHeight     =   210
      ScaleWidth      =   2100
      TabIndex        =   34
      Top             =   5760
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Menu Game 
      Caption         =   "游戏(&G)"
      Begin VB.Menu Begin 
         Caption         =   "开始游戏(&B)"
         Shortcut        =   {F2}
      End
      Begin VB.Menu Afresh 
         Caption         =   "重新开始(&A)"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu Stop 
         Caption         =   "暂停游戏(&I)"
         Shortcut        =   ^P
      End
      Begin VB.Menu BACDABD 
         Caption         =   "-"
      End
      Begin VB.Menu SaveM 
         Caption         =   "保存进度(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu SaveP 
         Caption         =   "保存图片(&P)"
      End
      Begin VB.Menu Open 
         Caption         =   "载入进度(&O)"
         Shortcut        =   ^O
      End
      Begin VB.Menu ABCDAB 
         Caption         =   "-"
      End
      Begin VB.Menu Option 
         Caption         =   "选择关卡(&O)"
         Shortcut        =   ^C
      End
      Begin VB.Menu ABDEFC 
         Caption         =   "-"
      End
      Begin VB.Menu Quit 
         Caption         =   "退出游戏(&Q)"
      End
   End
   Begin VB.Menu MnuRegister 
      Caption         =   "注册(&R)"
      Begin VB.Menu SetupTo 
         Caption         =   "本目录安装(&S)"
      End
      Begin VB.Menu MnuSetting 
         Caption         =   "设置(&T)"
      End
   End
   Begin VB.Menu Luyi 
      Caption         =   "录影(&L)"
      Visible         =   0   'False
      Begin VB.Menu Save 
         Caption         =   "保存在(&S)..."
      End
      Begin VB.Menu sasdds 
         Caption         =   "-"
      End
      Begin VB.Menu Bofa 
         Caption         =   "播放(&B)"
      End
      Begin VB.Menu Stop1 
         Caption         =   "停止(&T)"
      End
      Begin VB.Menu Luy 
         Caption         =   "录影(&L)"
      End
      Begin VB.Menu asede 
         Caption         =   "-"
      End
      Begin VB.Menu SaveAs 
         Caption         =   "保存当前场景(&A)"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "帮助(&H)"
      Begin VB.Menu About 
         Caption         =   "关于..."
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z
Const LEN_TANK = 28
Const LEN_BLOCK = 32
Const WG = 16
Const HWG = 8
Const HHWG = 4
Const TS_UP = 0
Const TS_DOWN = 2
Const TS_LEFT = 3
Const TS_RIGHT = 1
Const VD_EBULLET = 8
Const MAPZN = 6
'Const LVLMAPNAME As String = "市内巷战峰回路转森林野战古堡墟舍断壁残垣林密春深铜墙铁壁渡河远征钢铁盾牌怪物阵垒森林绿带水畔之战街心花园骷髅城堡巡回巷战空中花园隐密工厂王者魔杖大兵压境胜负之战"
Const strCAPTION As String = "坦克大战"
Dim GameState As Integer
Const GS_SPLASH = 0
Const GS_GAMING = 1
Const GS_GAMEOVER = 2
Const GS_RESULT = 3
Const GS_PAUSE = 4

Dim SleepType As Integer
Const GE_WIN = 1
Const GE_START = 2

Dim KingDie As Boolean
Dim PlayerNum As Integer
Dim Map(25, 25) As Byte
Dim Level As Integer
Dim LEVELID As String
Dim defLVL As Integer
Dim FAP As String
Dim BornNum(4) As Integer

Dim EBullet(19) As Bullet
Dim MinEBullet As Integer
Dim PBullet(7) As Bullet
Dim MinBullet(1) As Integer

Dim PLTank(1) As Tank
Dim PLMove(1) As Boolean
Dim PLLife(1) As Integer

Dim EYTank(3) As Tank
Dim EYMove(3) As Boolean
Dim MinTank As Integer
Dim Remain As Integer, NumPresent As Integer  'ii
Dim BornMovie(4) As Movie
Dim Explode(19) As Movie
Dim Bonus As BonusRC
Dim EatBonus(2) As Integer

Dim MinExplode As Integer
Dim Remark(1, 3) As Integer

Public App_Path As String
Dim cDib As New cDIBSection
Dim Setting As Configure
Dim Prevec As Boolean

Private Sub About_Click()
    frmAbout.Show
End Sub

Private Sub Afresh_Click()
    For i = 0 To 3
        EYTank(i).Show = False
    Next i
    For i = 0 To 1
        PLTank(i).Show = False
    Next i
    GameState = GS_SPLASH
    ShowSplash
    defLVL = 1
End Sub

Private Sub Begin_Click()
    If GameState = GS_SPLASH Then
        Start
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If GameState <> GS_SPLASH Then
        If MsgBox("您真的要退出游吗?", vbYesNo + vbInformation, "提示") = vbYes Then
            End
        Else
            Cancel = 1
        End If
    Else
        End
    End If
End Sub

Private Sub Open_Click()
    CoDi.Filter = "游戏进度文件(.gpf)|*.gpf"
    CoDi.ShowOpen
    If CoDi.FileName <> "" Then
        Open CoDi.FileName For Binary Access Read As 1
            Get #1, , Map
            Get #1, , PLTank
            Get #1, , PBullet
            Get #1, , EYTank
            Get #1, , EBullet
            Get #1, , Bonus
            Get #1, , Remain
            Get #1, , NumPresent
            Get #1, , GameState
            Get #1, , PLLife
        Close #1
    End If
End Sub

Private Sub Option_Click()
    Dim n As Integer
    n = FileLen("map.bin") \ SectionLen
    If GameState = GS_SPLASH Then
        defLVL = Val(InputBox("请输入关口序号(1-" & n & "):", "选关"))
        If defLVL > n Or defLVL < 1 Then defLVL = 1
        Start
    End If
End Sub

Private Sub Quit_Click()
    Unload Me
End Sub

Private Sub RunBonus_Timer()
    If GameState <> GS_GAMING And GameState <> GS_GAMEOVER Then Exit Sub
    Randomize Now
    Bonus.Show = False
    If Remain > 15 Then Exit Sub
    If 20 * Rnd > 8 Then Exit Sub
    With Bonus
        .T = Int(100 * Rnd Mod 7)
        .x = Int(1000 * Rnd Mod 390)
        .y = Int(1000 * Rnd Mod 390)
        .W = LEN_BLOCK
        .H = LEN_BLOCK
        .Show = True
    End With
End Sub

Private Sub BornKeeper_Timer()
    Dim flag As Boolean, tmp As Integer, i As Integer
    If GameState <> GS_GAMING And GameState <> GS_GAMEOVER Then Exit Sub
    Randomize Now 'Timer
    flag = False

    For i = 0 To 4
        If BornMovie(i).No > 0 Then
            BornMovie(i).Show = True
            flag = True
            If BornNum(i) = 1 Then BornMovie(i).No = BornMovie(i).No - 1 Else BornMovie(i).No = BornMovie(i).No + 1
            If BornMovie(i).No > 4 Or BornMovie(i).No < 2 Then BornNum(i) = BornNum(i) + 1
            Select Case i
            Case 0
                If BornNum(0) >= 3 Then
                    Do While EYTank(MinTank).Show
                        MinTank = (MinTank + 1) Mod 4
                    Loop
                    Remain = Remain - 1
                    With EYTank(MinTank)
                        .x = 0
                        .y = 0
                        .FANG = TS_DOWN
                        .FHP = 1
                        .TYPE = RndType
                        If .TYPE > 5 Then
                            .FER = 2
                        ElseIf .TYPE > 3 Then
                            .FHP = 4
                            .FER = 1
                        ElseIf .TYPE > 1 Then
                            .FER = 4
                        Else
                            .FER = 3
                        End If
                        .Show = True
                    End With
                    NumPresent = NumPresent + 1
                    Image3(Remain).Visible = False
                End If
            Case 1
                If BornNum(1) >= 3 Then
                    Do While EYTank(MinTank).Show
                        MinTank = (MinTank + 1) Mod 4
                    Loop
                    Remain = Remain - 1
                    With EYTank(MinTank)
                        .x = 192
                        .y = 0
                        .FANG = TS_DOWN
                        .FHP = 1
                        .TYPE = RndType
                        If .TYPE > 5 Then
                            .FER = 2
                        ElseIf .TYPE > 3 Then
                            .FHP = 4
                            .FER = 1
                        ElseIf .TYPE > 1 Then
                            .FER = 4
                        Else
                            .FER = 3
                        End If
                        .Show = True
                    End With
                    NumPresent = NumPresent + 1
                    Image3(Remain).Visible = False
                End If
            Case 2
                If BornNum(2) >= 3 Then
                    Do While EYTank(MinTank).Show
                        MinTank = (MinTank + 1) Mod 4
                    Loop
                    Remain = Remain - 1
                    With EYTank(MinTank)
                        .x = 384
                        .y = 0
                        .FANG = TS_DOWN
                        .FHP = 1
                        .TYPE = RndType
                        If .TYPE > 5 Then
                            .FER = 2
                        ElseIf .TYPE > 3 Then
                            .FHP = 4
                            .FER = 1 ' 2
                        ElseIf .TYPE > 1 Then
                            .FER = 4
                        Else
                            .FER = 3
                        End If
                        .Show = True
                    End With
                    NumPresent = NumPresent + 1
                    
                    Image3(Remain).Visible = False
                End If
            Case 3
                If BornNum(3) >= 3 Then
                    With PLTank(0)
                        .x = 128
                        .y = 384
                        .FANG = TS_UP
                        If .FHP < 1 Then .FHP = 1: .ProWa = False
                        If Not Prevec Then
                            .Shield = False
                            .FHP = 1
                        End If
                        .Show = True
                    End With
                End If
            Case 4
                If BornNum(4) >= 3 Then
                    With PLTank(1)
                        .x = 256
                        .y = 384
                        .FANG = TS_UP
                        If .FHP < 1 Then .FHP = 1: .ProWa = False
                        If Not Prevec Then
                            .Shield = False
                            .FHP = 1
                        End If
                        .Show = True
                    End With
                End If
            End Select
            If BornNum(i) = 3 Then BornMovie(i).No = 0: BornNum(i) = 0: BornMovie(i).Show = False
        End If
    Next i
    If Not flag Then
        If Remain And NumPresent < 4 Then
            BornMovie(Int(100 * Rnd) Mod 3).No = 1
        End If
    End If
    If NumPresent < 4 And Remain Then flag = True
    BornKeeper.Enabled = flag
End Sub

Private Sub Bullet_Timer()
    Dim i As Integer, xi As Integer, yi As Integer, xb As Integer, yb As Integer
    If GameState <> GS_GAMING And GameState <> GS_GAMEOVER Then Exit Sub
    For i = 0 To 19
        If EBullet(i).Show Then
            With EBullet(i)
                Select Case .FANG
                Case TS_UP
                    xi = .x \ WG: xb = .x - xi * WG
                    yi = .y \ WG: yb = .y - yi * WG
                    If yi > 0 Then
'                        Debug.Assert .TYPE = 0
                        If Map(xi, yi) > 5 Then '- 1 - 1
                            Map(xi, yi) = 0
                            .Show = False
                        End If
                        If Map(xi, yi - 1) = 1 Then
                            If .TYPE Then Map(xi, yi - 1) = 0 Else Map(xi, yi - 1) = MAPZN + .FANG
                            .Show = False
                        End If
                        If Map(xi, yi - 1) = 3 Then
                            If .TYPE Then Map(xi, yi - 1) = 0
                            sndPlaySound "HIT.wav", 1
                            .Show = False
                        End If
                        
                        If xb > HWG Then
                            If Map(xi + 1, yi) > 5 Then ' - 1 - 1
                                Map(xi + 1, yi) = 0
                                .Show = False
                            End If
                            If Map(xi + 1, yi - 1) = 1 Then
                                If .TYPE Then Map(xi + 1, yi - 1) = 0 Else Map(xi + 1, yi - 1) = MAPZN + .FANG
                                .Show = False
                            End If
                            If Map(xi + 1, yi - 1) = 3 Then
                                If .TYPE Then Map(xi + 1, yi - 1) = 0
                                sndPlaySound "HIT.wav", 1
                                .Show = False
                            End If
                        End If
                    End If
            
                    .y = .y - VD_EBULLET
                    If .y < 0 Then .Show = False
                Case TS_DOWN
                    xi = .x \ WG: xb = .x - xi * WG
                    yi = (.y + HWG) \ WG: yb = .y + HWG - yi * WG
                    If yi < 26 Then
                        If Map(xi, yi) > 5 Then
                            Map(xi, yi) = 0
                            .Show = False
                        End If
                        If Map(xi, yi) = 1 Then
                            If .TYPE Then Map(xi, yi) = 0 Else Map(xi, yi) = MAPZN + .FANG
                            .Show = False
                        End If
                        If Map(xi, yi) = 3 Then
                            If .TYPE Then Map(xi, yi) = 0
                            sndPlaySound "HIT.wav", 1
                            .Show = False
                        End If
                        
                        If xb > HWG Then
                            If Map(xi + 1, yi) > 5 Then
                                Map(xi + 1, yi) = 0
                                .Show = False
                            End If
                            If Map(xi + 1, yi) = 1 Then
                                If .TYPE Then Map(xi + 1, yi) = 0 Else Map(xi + 1, yi) = MAPZN + .FANG
                                .Show = False
                            End If
                            If Map(xi + 1, yi) = 3 Then
                                If .TYPE Then Map(xi + 1, yi) = 0
                                sndPlaySound "HIT.wav", 1
                                .Show = False
                            End If
                        End If
                        If yi = 24 And (xi = 12 Or xi = 13) Then
                            BitBlt PicBack.hdc, 12 * WG, 24 * WG, LEN_BLOCK, LEN_BLOCK, PicIcon(0).hdc, 0, 0, SRCCOPY
                            GameOver KingDie
                            KingDie = True
                        End If
                    End If
                    .y = .y + VD_EBULLET
                    If .y > 420 Then .Show = False
                Case TS_LEFT
                    xi = .x \ WG: xb = .x - xi * WG
                    yi = .y \ WG: yb = .y - yi * WG
                    If xi > 0 Then
                        If Map(xi, yi) > 5 Then ' - 1 - 1
                            Map(xi, yi) = 0
                            .Show = False
                        End If
                        If Map(xi - 1, yi) = 1 Then
                            If .TYPE Then Map(xi - 1, yi) = 0 Else Map(xi - 1, yi) = MAPZN + .FANG
                            'Map(xi - 1, yi) = 0
                            .Show = False
                        End If
                        If Map(xi - 1, yi) = 3 Then
                            If .TYPE Then Map(xi - 1, yi) = 0
                            sndPlaySound "HIT.wav", 1
                            .Show = False
                        End If
                        
                        If yb > HWG Then
                            If Map(xi, yi + 1) > 5 Then '- 1   - 1
                                Map(xi, yi + 1) = 0
                                .Show = False
                            End If
                            If Map(xi - 1, yi + 1) = 1 Then
                                If .TYPE Then Map(xi - 1, yi + 1) = 0 Else Map(xi - 1, yi + 1) = MAPZN + .FANG
                                'Map(xi - 1, yi + 1) = 0
                                .Show = False
                            End If
                            If Map(xi - 1, yi + 1) = 3 Then
                                If .TYPE Then Map(xi - 1, yi + 1) = 0
                                sndPlaySound "HIT.wav", 1
                                .Show = False
                            End If
                        End If
                        
                        If xi = 13 And (yi = 24 Or yi = 25) Then
                            BitBlt PicBack.hdc, 12 * WG, 24 * WG, LEN_BLOCK, LEN_BLOCK, PicIcon(0).hdc, 0, 0, SRCCOPY
                            GameOver KingDie
                            KingDie = True
                        End If
                    End If
                    .x = .x - VD_EBULLET
                    If .x < 0 Then .Show = False
                Case TS_RIGHT
                    xi = (.x + HWG) \ WG: xb = .x + HWG - xi * WG
                    yi = .y \ WG: yb = .y - yi * WG
                    If xi < 26 Then
                        If Map(xi, yi) > 5 Then
                            Map(xi, yi) = 0
                            .Show = False
                        End If
                        If Map(xi, yi) = 1 Then
                            If .TYPE Then Map(xi, yi) = 0 Else Map(xi, yi) = MAPZN + .FANG
                            .Show = False
                        End If

                        
                        If Map(xi, yi) = 3 Then
                            If .TYPE Then Map(xi, yi) = 0
                            sndPlaySound "HIT.wav", 1
                            .Show = False
                        End If
                        
                        If yb > HWG Then
                            If Map(xi, yi + 1) > 5 Then
                                Map(xi, yi + 1) = 0
                                .Show = False
                            End If
                            If Map(xi, yi + 1) = 1 Then
                                If .TYPE Then Map(xi, yi + 1) = 0 Else Map(xi, yi + 1) = MAPZN + .FANG
                                .Show = False
                            End If

                            If Map(xi, yi + 1) = 3 Then
                                If .TYPE Then Map(xi, yi + 1) = 0
                                sndPlaySound "HIT.wav", 1
                                .Show = False
                            End If
                        End If
                        
                        If xi = 12 And (yi = 24 Or yi = 25) Then
                            BitBlt PicBack.hdc, 12 * WG, 24 * WG, LEN_BLOCK, LEN_BLOCK, PicIcon(0).hdc, 0, 0, SRCCOPY
                            GameOver KingDie
                            KingDie = True
                        End If
                    End If
                    .x = .x + VD_EBULLET
                    If .x > 420 Then .Show = False
                End Select

                For j = 0 To 1
                    If PLTank(j).Show Then
                        If PLTank(j).y <= .y + IIf(.TYPE, WG, HWG) And .y <= PLTank(j).y + LEN_TANK Then
                            If .x >= PLTank(j).x - IIf(.TYPE, WG, HWG) And .x <= PLTank(j).x + LEN_TANK Then
                                If PLTank(j).Shield Then
                                    If .TYPE Then PLTank(j).Shield = False
                                    .Show = False
                                    sndPlaySound "HIT.wav", 1
                                Else
                                    PLTank(j).FHP = PLTank(j).FHP - IIf(.TYPE, 2, 1)
                                    .Show = False
                                    If PLTank(j).FHP <= 0 Then
                                        PLTank(j).Show = False
                                        If PLLife(j) >= 0 Then
                                            PLLife(j) = PLLife(j) - 1
                                            PLLife(j) = IIf(PLLife(j) > 0, PLLife(j), 0)
                                            If PLLife(0) + PLLife(1) <= 0 Then
                                                If PLTank(0).Show = False And ((PLTank(1).Show = False And PlayerNum = 1) Or PlayerNum = 0) Then GameOver 1
                                            ElseIf PLLife(j) > 0 Then
                                                BornMovie(IIf(j, 4, 3)).No = 1
                                                BornKeeper.Enabled = True
                                            End If
                                        End If
                                        Explode(MinExplode).Show = True
                                        Explode(MinExplode).No = 0
                                        Explode(MinExplode).NUM = 6
                                        Explode(MinExplode).x = PLTank(j).x - HWG
                                        Explode(MinExplode).y = PLTank(j).y - HWG
                                        MinExplode = (MinExplode + 1) Mod 20
                                        sndPlaySound "BANG.wav", 1
                                    ElseIf PLTank(j).FHP < 3 Then
                                        PLTank(j).ProWa = False
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next j
               ' For j = 0 To 7
               '     If PBullet(j).Y <= .Y + IIf(.TYPE, WG, HWG) And .Y <= PBullet(j).Y + IIf(PBullet(j).TYPE, WG, HWG) Then
               '         If .X >= PBullet(j).X - IIf(.TYPE, WG, HWG) And .X <= PBullet(j).X + IIf(PBullet(j).TYPE, WG, HWG) Then
               '             .Show = False
               '             PBullet(j).Show = False
               '         End If
               '     End If
               ' Next j
            End With
        End If
    Next i
    
    For i = 0 To 7
        If PBullet(i).Show Then
            With PBullet(i)
                Select Case .FANG
                Case TS_UP
                    xi = .x \ WG: xb = .x - xi * WG
                    yi = .y \ WG: yb = .y - yi * WG
                    If yi > 0 Then
                        If Map(xi, yi - 1) > 5 Then
                            Map(xi, yi - 1) = 0
                            .Show = False
                        End If
                        If Map(xi, yi - 1) = 1 Then
                            If .TYPE Then Map(xi, yi - 1) = 0 Else Map(xi, yi - 1) = MAPZN + .FANG
                            .Show = False
                        End If
                        If Map(xi, yi - 1) = 3 Then
                            If .TYPE Then Map(xi, yi - 1) = 0
                            sndPlaySound "HIT.wav", 1
                            .Show = False
                        End If
                        
                        If xb > HWG Then
                            If Map(xi + 1, yi - 1) > 5 Then
                                Map(xi + 1, yi - 1) = 0
                                .Show = False
                            End If
                            If Map(xi + 1, yi - 1) = 1 Then
                                If .TYPE Then Map(xi + 1, yi - 1) = 0 Else Map(xi + 1, yi - 1) = MAPZN + .FANG
                                .Show = False
                            End If
                            If Map(xi + 1, yi - 1) = 3 Then
                                If .TYPE Then Map(xi + 1, yi - 1) = 0
                                sndPlaySound "HIT.wav", 1
                                .Show = False
                            End If
                        End If
                    End If
            
                    .y = .y - VD_EBULLET
                    If .y < 0 Then .Show = False
                Case TS_DOWN
                    xi = .x \ WG: xb = .x - xi * WG
                    yi = (.y + HWG) \ WG: yb = .y + HWG - yi * WG
                    If yi < 26 Then
                        If Map(xi, yi) > 5 Then
                            Map(xi, yi) = 0
                            .Show = False
                        End If
                        If Map(xi, yi) = 1 Then
                            If .TYPE Then Map(xi, yi) = 0 Else Map(xi, yi) = MAPZN + .FANG
                            .Show = False
                        End If
                        If Map(xi, yi) = 3 Then
                            If .TYPE Then Map(xi, yi) = 0
                            sndPlaySound "HIT.wav", 1
                            .Show = False
                        End If
                        
                        If xb > HWG Then
                            If Map(xi + 1, yi) > 5 Then
                                Map(xi + 1, yi) = 0
                                .Show = False
                            End If
                            If Map(xi + 1, yi) = 1 Then
                                If .TYPE Then Map(xi + 1, yi) = 0 Else Map(xi + 1, yi) = MAPZN + .FANG
                                .Show = False
                            End If
                            If Map(xi + 1, yi) = 3 Then
                                If .TYPE Then Map(xi + 1, yi) = 0
                                sndPlaySound "HIT.wav", 1
                                .Show = False
                            End If
                        End If
                        If yi = 24 And (xi = 12 Or xi = 13) Then
                            BitBlt PicBack.hdc, 12 * WG, 24 * WG, LEN_BLOCK, LEN_BLOCK, PicIcon(0).hdc, 0, 0, SRCCOPY
                            GameOver KingDie
                            KingDie = True
                        End If
                    End If
                    .y = .y + VD_EBULLET
                    If .y > 420 Then .Show = False
                Case TS_LEFT
                    xi = .x \ WG: xb = .x - xi * WG
                    yi = .y \ WG: yb = .y - yi * WG
                    If xi > 0 Then
                        If Map(xi - 1, yi) > 5 Then
                            Map(xi - 1, yi) = 0
                            .Show = False
                        End If
                        If Map(xi - 1, yi) = 1 Then
                            If .TYPE Then Map(xi - 1, yi) = 0 Else Map(xi - 1, yi) = MAPZN + .FANG
                            'Map(xi - 1, yi) = 0
                            .Show = False
                        End If
                        If Map(xi - 1, yi) = 3 Then
                            If .TYPE Then Map(xi - 1, yi) = 0
                            sndPlaySound "HIT.wav", 1
                            .Show = False
                        End If
                        
                        If yb > HWG Then
                            If Map(xi - 1, yi + 1) > 5 Then
                                Map(xi - 1, yi + 1) = 0
                                .Show = False
                            End If
                            If Map(xi - 1, yi + 1) = 1 Then
                                If .TYPE Then Map(xi - 1, yi + 1) = 0 Else Map(xi - 1, yi + 1) = MAPZN + .FANG
                                'Map(xi - 1, yi + 1) = 0
                                .Show = False
                            End If
                            If Map(xi - 1, yi + 1) = 3 Then
                                If .TYPE Then Map(xi - 1, yi + 1) = 0
                                sndPlaySound "HIT.wav", 1
                                .Show = False
                            End If
                        End If
                        
                        If xi = 13 And (yi = 24 Or yi = 25) Then
                            BitBlt PicBack.hdc, 12 * WG, 24 * WG, LEN_BLOCK, LEN_BLOCK, PicIcon(0).hdc, 0, 0, SRCCOPY
                            GameOver KingDie
                            KingDie = True
                        End If
                    End If
                    .x = .x - VD_EBULLET
                    If .x < 0 Then .Show = False
                Case TS_RIGHT
                    xi = (.x + HWG) \ WG: xb = .x + HWG - xi * WG
                    yi = .y \ WG: yb = .y - yi * WG
                    If xi < 26 Then
                        If Map(xi, yi) > 5 Then
                            Map(xi, yi) = 0
                            .Show = False
                        End If
                        If Map(xi, yi) = 1 Then
                            If .TYPE Then Map(xi, yi) = 0 Else Map(xi, yi) = MAPZN + .FANG
                            .Show = False
                        End If

                        
                        If Map(xi, yi) = 3 Then
                            If .TYPE Then Map(xi, yi) = 0
                            sndPlaySound "HIT.wav", 1
                            .Show = False
                        End If
                        
                        If yb > HWG Then
                            If Map(xi, yi + 1) > 5 Then
                                Map(xi, yi + 1) = 0
                                .Show = False
                            End If
                            If Map(xi, yi + 1) = 1 Then
                                If .TYPE Then Map(xi, yi + 1) = 0 Else Map(xi, yi + 1) = MAPZN + .FANG
                                .Show = False
                            End If

                            If Map(xi, yi + 1) = 3 Then
                                If .TYPE Then Map(xi, yi + 1) = 0
                                sndPlaySound "HIT.wav", 1
                                .Show = False
                            End If
                        End If
                        
                        If xi = 12 And (yi = 24 Or yi = 25) Then
                            BitBlt PicBack.hdc, 12 * WG, 24 * WG, LEN_BLOCK, LEN_BLOCK, PicIcon(0).hdc, 0, 0, SRCCOPY
                            GameOver KingDie
                            KingDie = True
                        End If
                    End If
                    .x = .x + VD_EBULLET
                    If .x > 420 Then .Show = False
                End Select
                For j = 0 To 3
                    If EYTank(j).Show Then
                        If EYTank(j).y <= .y + IIf(.TYPE, WG, HWG) And .y <= EYTank(j).y + LEN_TANK Then
                            If .x >= EYTank(j).x - IIf(.TYPE, WG, HWG) And .x <= EYTank(j).x + LEN_TANK Then
                                If EYTank(j).Shield Then
                                    If .TYPE Then EYTank(j).Shield = False
                                Else
                                    EYTank(j).FHP = EYTank(j).FHP - IIf(.TYPE, 2, 1)
                                    .Show = False
                                    If EYTank(j).FHP <= 0 Then
                                        Remark(IIf(i > 3, 1, 0), EYTank(j).TYPE \ 2) = Remark(IIf(i > 3, 1, 0), EYTank(j).TYPE \ 2) + 1
                                        EYTank(j).Show = False
                                        NumPresent = NumPresent - 1
                                        Explode(MinExplode).Show = True
                                        Explode(MinExplode).No = 0
                                        Explode(MinExplode).NUM = 6
                                        Explode(MinExplode).x = EYTank(j).x - HWG
                                        Explode(MinExplode).y = EYTank(j).y - HWG
                                        MinExplode = (MinExplode + 1) Mod 20
                                        If NumPresent = 0 Then
                                            If Remain <= 0 Then
                                                Sleep GE_WIN, 2
                                            End If
                                        ElseIf NumPresent < 4 Then
                                            If Remain Then
                                                BornKeeper.Enabled = True
                                            End If
                                        End If
                                    End If
                                End If
                                sndPlaySound "BANG.wav", 1
                            End If
                        End If
                    End If
                Next j
            End With
        End If
    Next i
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim xi As Integer, yi As Integer, xb As Integer, yb As Integer, Tndex As Integer
    Dim xc As Integer, yc As Integer, Tmpint As Integer, Slip As Boolean
    If GameState = GS_SPLASH Then
        Select Case KeyCode
        Case vbKeyDown
            PlayerNum = (PlayerNum + 1) Mod 2
            ShowSplash
        Case vbKeyUp
            PlayerNum = Abs(PlayerNum - 1) Mod 2
            ShowSplash
        Case vbKey1
            PlayerNum = 0
            ShowSplash
        Case vbKey2
            PlayerNum = 1
            ShowSplash
        Case vbKeyReturn
            Start '5
        End Select
    ElseIf GameState = GS_GAMING Then
        For Tndex = 0 To 1
            If PLTank(Tndex).Show Then
                With PLTank(Tndex)
                    xc = .x \ WG: yc = .y \ WG
                    If GetAsyncKeyState(Setting.Player_DOWN(Tndex)) < 0 Then
                        If .FANG = TS_DOWN Then
                            xi = xc: xb = .x - xi * WG
                            yi = (.y + LEN_TANK) \ WG: yb = .y + LEN_TANK - yi * WG
                            If yb = 0 And yi < 26 Then
                                If Map(xi, yi) = 1 Or Map(xi, yi) = 3 Or Map(xi, yi) = 4 Then GoTo LT
                                If Map(xi, yi) = MAPZN + TS_LEFT And xb < HWG Then GoTo LT
                                If Map(xi, yi) = MAPZN + TS_RIGHT Or Map(xi, yi) = MAPZN + TS_UP Then GoTo LT

                                If Map(xi + 1, yi) = 1 Or Map(xi + 1, yi) = 3 Or Map(xi + 1, yi) = 4 Or _
                                    Map(xi + 1, yi) = MAPZN + TS_RIGHT Or Map(xi + 1, yi) = MAPZN + TS_LEFT _
                                    Or Map(xi + 1, yi) = MAPZN + TS_UP Then GoTo LT
                                                            
                                If xi < 24 Then If xb > 4 And (Map(xi + 2, yi) = 1 Or Map(xi + 2, yi) = 3 Or Map(xi + 2, yi) = 4 Or Map(xi + 2, yi) = MAPZN + TS_UP Or Map(xi + 2, yi) = MAPZN + TS_LEFT) Then GoTo LT
                            End If
                            If yb = HWG And yi < 26 Then
                                If Map(xi, yi) = MAPZN + TS_DOWN Then GoTo LT
                                If Map(xi + 1, yi) = MAPZN + TS_DOWN Then GoTo LT
                                If xi < 24 Then If xb > HWG And Map(xi + 2, yi) = MAPZN + TS_DOWN Then GoTo LT
                            End If
                            For i = 0 To 3
                                If EYTank(i).Show Then
                                    If EYTank(i).y - .y < 4 + LEN_TANK And .y < EYTank(i).y _
                                        And .x > EYTank(i).x - LEN_TANK And .x < EYTank(i).x + LEN_TANK Then GoTo LT
                                End If
                            Next i
                            If PLTank(1 - Tndex).Show Then
                                If PLTank(1 - Tndex).y - .y < 4 + LEN_TANK And .y < PLTank(1 - Tndex).y _
                                    And .x > PLTank(1 - Tndex).x - LEN_TANK And .x < PLTank(1 - Tndex).x + LEN_TANK Then GoTo LT
                            End If
                            If .y < 388 Then .y = .y + 4
                            .Move = True
                        Else
                            .FANG = TS_DOWN
                        End If
                    ElseIf GetAsyncKeyState(Setting.Player_UP(Tndex)) < 0 Then
                        If .FANG = TS_UP Then
                            xi = xc: xb = .x - xi * WG
                            yi = yc: yb = .y - yi * WG
                            If yb = 0 And yi > 0 Then
                                If Map(xi, yi - 1) = 1 Or Map(xi, yi - 1) = 3 Or Map(xi, yi - 1) = 4 Then GoTo LT
                                If Map(xi, yi - 1) = MAPZN + TS_LEFT And xb < HWG Then GoTo LT
                                If Map(xi, yi - 1) = MAPZN + TS_RIGHT Or Map(xi, yi - 1) = MAPZN + TS_DOWN Then GoTo LT
                                
                                If Map(xi + 1, yi - 1) = 1 Or Map(xi + 1, yi - 1) = 3 Or Map(xi + 1, yi - 1) = 4 Or _
                                    Map(xi + 1, yi - 1) = MAPZN + TS_RIGHT Or Map(xi + 1, yi - 1) = MAPZN + TS_LEFT _
                                    Or Map(xi + 1, yi - 1) = MAPZN + TS_DOWN Then GoTo LT
                                
                                If xi < 24 Then If xb > 4 And (Map(xi + 2, yi - 1) = 1 Or Map(xi + 2, yi - 1) = 3 Or Map(xi + 2, yi - 1) = 4 Or Map(xi + 2, yi - 1) = MAPZN + TS_DOWN Or Map(xi + 2, yi - 1) = MAPZN + TS_LEFT) Then GoTo LT
                            End If
                            If yb = HWG And yi > 0 Then
                                If Map(xi, yi) = MAPZN + TS_UP Then GoTo LT
                                If Map(xi + 1, yi) = MAPZN + TS_UP Then GoTo LT
                                If xi < 24 Then If xb > HWG And Map(xi + 2, yi) = MAPZN + TS_UP Then GoTo LT
                            End If
                            For i = 0 To 3
                                If EYTank(i).Show Then
                                    If .y - EYTank(i).y < 4 + LEN_TANK And .y > EYTank(i).y _
                                        And .x > EYTank(i).x - LEN_TANK And .x < EYTank(i).x + LEN_TANK Then GoTo LT
                                End If
                            Next i
                            If PLTank(1 - Tndex).Show Then
                                If .y - PLTank(1 - Tndex).y < 4 + LEN_TANK And .y > PLTank(1 - Tndex).y _
                                    And .x > PLTank(1 - Tndex).x - LEN_TANK And .x < PLTank(1 - Tndex).x + LEN_TANK Then GoTo LT
                            End If
                            If .y > 3 Then .y = .y - 4
                            .Move = True
                        Else
                            .FANG = TS_UP
                        End If
                    ElseIf GetAsyncKeyState(Setting.Player_LEFT(Tndex)) < 0 Then
                        If .FANG = TS_LEFT Then
                            xi = xc: xb = .x - xi * WG
                            yi = yc: yb = .y - yi * WG
                            If xb = 0 And xi > 0 Then
                                If Map(xi - 1, yi) = 1 Or Map(xi - 1, yi) = 3 Or Map(xi - 1, yi) = 4 Then GoTo LT
                                If Map(xi - 1, yi) = MAPZN + TS_UP And yb < HWG Then GoTo LT
                                If Map(xi - 1, yi) = MAPZN + TS_RIGHT Or Map(xi - 1, yi) = MAPZN + TS_DOWN Then GoTo LT
                                
                                If Map(xi - 1, yi + 1) = 1 Or Map(xi - 1, yi + 1) = 3 Or Map(xi - 1, yi + 1) = 4 Or _
                                    Map(xi - 1, yi + 1) = MAPZN + TS_RIGHT Or Map(xi - 1, yi + 1) = MAPZN + TS_UP _
                                    Or Map(xi - 1, yi + 1) = MAPZN + TS_DOWN Then GoTo LT
                                
                                If yi < 24 Then If yb > 4 And (Map(xi - 1, yi + 2) = 1 Or Map(xi - 1, yi + 2) = 3 Or Map(xi - 1, yi + 2) = 4 Or Map(xi - 1, yi + 2) = MAPZN + TS_UP Or Map(xi - 1, yi + 2) = MAPZN + TS_RIGHT) Then GoTo LT
                            End If
                            If xb = HWG And xi > 0 Then
                                If Map(xi, yi) = MAPZN + TS_LEFT Then GoTo LT
                                If Map(xi, yi + 1) = MAPZN + TS_LEFT Then GoTo LT
                                If yi < 24 Then If yb > HWG And Map(xi, yi + 2) = MAPZN + TS_LEFT Then GoTo LT
                            End If
                            For i = 0 To 3
                                If EYTank(i).Show Then
                                    If EYTank(i).y <= .y + LEN_TANK And .y <= EYTank(i).y + LEN_TANK _
                                        And .x - EYTank(i).x < 4 + LEN_TANK And .x > EYTank(i).x Then GoTo LT
                                End If
                            Next i
                            If PLTank(1 - Tndex).Show Then
                                If PLTank(1 - Tndex).y <= .y + LEN_TANK And .y <= PLTank(1 - Tndex).y + LEN_TANK _
                                    And .x - PLTank(1 - Tndex).x < 4 + LEN_TANK And .x > PLTank(1 - Tndex).x Then GoTo LT
                            End If
                            If .x > 3 Then .x = .x - 4
                            .Move = True
                        Else
                            .FANG = TS_LEFT
                        End If
                    ElseIf GetAsyncKeyState(Setting.Player_RIGHT(Tndex)) < 0 Then
                        If .FANG = TS_RIGHT Then
                            xi = (.x + LEN_TANK) \ WG: xb = .x + LEN_TANK - xi * WG
                            yi = .y \ WG: yb = .y - yi * WG
                            If xb = 0 And xi < 26 Then
                                If Map(xi, yi) = 1 Or Map(xi, yi) = 3 Or Map(xi, yi) = 4 Then GoTo LT
                                If Map(xi, yi) = MAPZN + TS_UP And yb < HWG Then GoTo LT
                                If Map(xi, yi) = MAPZN + TS_LEFT Or Map(xi - 1, yi) = MAPZN + TS_DOWN Then GoTo LT
                                
                                If Map(xi, yi + 1) = 1 Or Map(xi, yi + 1) = 3 Or Map(xi, yi + 1) = 4 Or _
                                    Map(xi, yi + 1) = MAPZN + TS_LEFT Or Map(xi, yi + 1) = MAPZN + TS_UP _
                                    Or Map(xi, yi + 1) = MAPZN + TS_DOWN Then GoTo LT
                                
                                If yi < 24 Then If yb > 4 And (Map(xi, yi + 2) = 1 Or Map(xi, yi + 2) = 3 Or Map(xi, yi + 2) = 4 Or Map(xi, yi + 2) = MAPZN + TS_UP Or Map(xi, yi + 2) = MAPZN + TS_LEFT) Then GoTo LT
                            End If
                            If xb = HWG And xi < 26 Then
                                If Map(xi, yi) = MAPZN + TS_RIGHT Then GoTo LT
                                If Map(xi, yi + 1) = MAPZN + TS_RIGHT Then GoTo LT
                                If yi < 24 Then If yb > HWG And Map(xi, yi + 2) = MAPZN + TS_RIGHT Then GoTo LT
                            End If
                            For i = 0 To 3
                                If EYTank(i).Show Then
                                    If EYTank(i).y <= .y + LEN_TANK And .y <= EYTank(i).y + LEN_TANK _
                                        And EYTank(i).x - .x < 4 + LEN_TANK And .x < EYTank(i).x Then GoTo LT
                                End If
                            Next i
                            If PLTank(1 - Tndex).Show Then
                                If PLTank(1 - Tndex).y <= .y + LEN_TANK And .y <= PLTank(1 - Tndex).y + LEN_TANK _
                                    And PLTank(1 - Tndex).x - .x < 4 + LEN_TANK And .x < PLTank(1 - Tndex).x Then GoTo LT
                            End If
                            If .x < 388 Then .x = .x + 4
                            .Move = True
                        Else
                            .FANG = TS_RIGHT
                        End If
                    End If
                    
                    If Bonus.Show Then
                        If Bonus.y <= .y + LEN_BLOCK And .y <= Bonus.y + LEN_BLOCK Then
                            If .x >= Bonus.x - LEN_BLOCK And .x <= Bonus.x + LEN_BLOCK Then
                                Bonus.Show = False
                                EchoBonus Bonus.T, Tndex
                            End If
                        End If
                    End If
                    
                    If GetAsyncKeyState(Setting.Player_SHOOT(Tndex)) < 0 Then
                        Tmpint = 0
                        While PBullet(MinBullet(Tndex)).Show And Tmpint < 4
                            MinBullet(Tndex) = (MinBullet(Tndex) + 1) Mod 4 + IIf(Tndex, 4, 0)
                            Tmpint = Tmpint + 1
                        Wend

                        If Not PBullet(MinBullet(Tndex)).Show Then
                            If Not .ProWa Then
                                PBullet(MinBullet(Tndex)).FANG = .FANG
                                PBullet(MinBullet(Tndex)).Show = True
                                PBullet(MinBullet(Tndex)).x = .x + 10
                                PBullet(MinBullet(Tndex)).y = .y + 10
                                PBullet(MinBullet(Tndex)).TYPE = 0
                                MinBullet(Tndex) = (MinBullet(Tndex) + 1) Mod 4 + IIf(Tndex, 4, 0)
                                sndPlaySound "FIRE.wav", 1
    
                            Else
                                PBullet(MinBullet(Tndex)).FANG = .FANG
                                PBullet(MinBullet(Tndex)).Show = True
                                PBullet(MinBullet(Tndex)).x = .x + 10
                                PBullet(MinBullet(Tndex)).y = .y + 10
                                PBullet(MinBullet(Tndex)).TYPE = 1
                                MinBullet(Tndex) = (MinBullet(Tndex) + 1) Mod 4 + IIf(Tndex, 4, 0)
                                sndPlaySound "FIRE.wav", 17
                            End If
                        End If
                    End If
                End With
            Else
                If GetAsyncKeyState(Setting.Player_SHOOT(Tndex)) < 0 Then
                    If PLLife(1 - Tndex) > 1 And BornMovie(3 + Tndex).Show = False And PLLife(Tndex) = 0 Then
                        PLLife(1 - Tndex) = PLLife(1 - Tndex) - 1
                        PLLife(Tndex) = PLLife(Tndex) + 1
                        BornMovie(3 + Tndex).No = 1
                        BornKeeper.Enabled = True
                    End If
                End If
            End If
LT:
        Next Tndex
        PaintScreen
    End If
End Sub

Private Sub Form_Load()
    Open "Tank.cfg" For Binary Access Read As 1
    Get #1, , Setting
    Close #1
    PicBack.Move 8, 8
    ShowSplash
    defLVL = 1
    FAP = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\")
End Sub

Sub ShowSplash()
    PicBack.Cls
    BitBlt PicBack.hdc, 20, 100, 376, 222, PicSplash.hdc, 0, 0, SRCCOPY
    If PlayerNum Then
        BitBlt PicBack.hdc, 100, 300, LEN_TANK, LEN_TANK, PicMytank(0).hdc, 0, LEN_TANK * TS_RIGHT, SRCCOPY
    Else
        BitBlt PicBack.hdc, 100, 266, LEN_TANK, LEN_TANK, PicMytank(0).hdc, 0, LEN_TANK * TS_RIGHT, SRCCOPY
    End If
End Sub

Public Sub Read()
    Dim i As Integer
    Open "Map.bin" For Binary Access Read As #1 Len = SectionLen  '676
        Get #1, (Level - 1) * SectionLen + 1, Map
    Close
    
    i = 0
    Open "gninfo.txt" For Input As 1
        Do Until EOF(1) Or i = Level
            Line Input #1, LEVELID
            i = i + 1
        Loop
        If i <> Level Then LEVELID = "未命名关"
    Close
End Sub

Public Sub Start(Optional ByVal lvl As Integer)
    If lvl = 0 Then
        Level = defLVL
        PLLife(0) = 3
        If PlayerNum = 1 Then
            BornMovie(4).No = 1
            BornMovie(4).x = 256
            BornMovie(4).y = 384
            PLLife(1) = 3
        End If
    Else
        PLTank(1).Show = False
        PLTank(0).Show = False
        Bonus.Show = False
        Level = lvl
    End If
    
    lblLevel.Caption = Level
    For i = 0 To 7
        PBullet(i).Show = False
    Next i
    For i = 0 To 3
        Remark(0, i) = 0
        Remark(1, i) = 0
    Next i
    For i = 0 To 19
        Image3(i).Visible = True
        EBullet(i).Show = False
    Next i
    
    GameState = GS_GAMING
    Read
    Me.Caption = strCAPTION & "---任务" & Level & "---" & LEVELID
    Map(13, 24) = 4
    Map(12, 24) = 4
    PaintScreen
    Remain = 20
    NumPresent = 0
    
    BornMovie(0).No = 1
    BornMovie(0).x = 0
    BornMovie(0).y = 0
    BornMovie(1).No = 1
    BornMovie(1).x = 192
    BornMovie(1).y = 0
    BornMovie(2).No = 1
    BornMovie(2).x = 384
    BornMovie(2).y = 0
    
    If PLLife(0) Then BornMovie(3).No = 1
    BornMovie(3).x = 128
    BornMovie(3).y = 384
    If PlayerNum = 1 And PLLife(1) Then BornMovie(4).No = 1
    
    BornKeeper.Enabled = True
    MinBullet(1) = 4
    KingDie = False
End Sub

Public Sub PaintScreen()
    PicBack.Cls
    For i = 0 To 25
        For j = 0 To 25
            If Map(i, j) > 0 And Map(i, j) < 11 Then
                If Map(i, j) <> 2 Then
                    If Map(i, j) < 6 Then
                        If Map(i, j) = 4 Then
                            BitBlt PicBack.hdc, i * WG, j * WG, WG, WG, PicDX(IIf(Timer Mod 2, 3, 5)).hdc, 0, 0, vbSrcCopy
                        Else
                            BitBlt PicBack.hdc, i * WG, j * WG, WG, WG, PicDX(Map(i, j) - 1).hdc, 0, 0, vbSrcCopy
                        End If
                    Else
                        Select Case Map(i, j) - 6
                        Case TS_DOWN
                            BitBlt PicBack.hdc, i * WG, j * WG + HWG, WG, HWG, PicDX(0).hdc, 0, HWG, vbSrcCopy
                        Case TS_UP
                            BitBlt PicBack.hdc, i * WG, j * WG, WG, HWG, PicDX(0).hdc, 0, 0, vbSrcCopy
                        Case TS_LEFT
                            BitBlt PicBack.hdc, i * WG, j * WG, HWG, WG, PicDX(0).hdc, 0, 0, vbSrcCopy
                        Case TS_RIGHT
                            BitBlt PicBack.hdc, i * WG + HWG, j * WG, HWG, WG, PicDX(0).hdc, HWG, 0, vbSrcCopy
                        End Select
                    End If
                End If
            End If
        Next j
    Next i
    
    For i = 0 To 19
        If EBullet(i).Show Then
            If EBullet(i).TYPE Then
                TransparentPaint PicBack, PicBZ(1), EBullet(i).x, EBullet(i).y, HWG, HWG, HWG * EBullet(i).FANG, 0
            Else
                TransparentPaint PicBack, PicBZ(0), EBullet(i).x, EBullet(i).y, HWG, HWG, HWG * EBullet(i).FANG, 0
            End If
        End If
    Next i
    
    If KingDie Then
        BitBlt PicBack.hdc, 12 * WG, 24 * WG, LEN_BLOCK, LEN_BLOCK, PicIcon(0).hdc, 0, 0, SRCCOPY
    Else
        BitBlt PicBack.hdc, 192, 384, LEN_BLOCK, LEN_BLOCK, PicIcon(1).hdc, 0, 0, SRCCOPY
    End If
    
    For i = 0 To 7
        If PBullet(i).Show Then
            If PBullet(i).TYPE Then
                TransparentPaint PicBack, PicBZ(1).Picture, PBullet(i).x, PBullet(i).y, HWG, HWG, HWG * PBullet(i).FANG, 0
            Else
                TransparentPaint PicBack, PicBZ(0).Picture, PBullet(i).x, PBullet(i).y, HWG, HWG, HWG * PBullet(i).FANG, 0
            End If
        End If
    Next i
    
    For i = 0 To 3
        If EYTank(i).Show Then
            PaintETank i
        End If
    Next i
    
    For i = 0 To 1
        If PLTank(i).Show Then
            With PLTank(i)
                TransparentPaint PicBack, PicMytank(i).Picture, .x, .y, LEN_TANK, LEN_TANK, (.FHP - 1) * 2 * LEN_TANK + IIf(.Move, LEN_TANK, 0), .FANG * LEN_TANK, 0
                'BitBlt PicBack.hdc, .x, .y, LEN_TANK, LEN_TANK, PicMytank(i).hdc, (.FHP - 1) * 2 * LEN_TANK + IIf(.Move, LEN_TANK, 0), .FANG * LEN_TANK, vbSrcCopy
                .Move = False
                If .Shield Then
                    TransparentPaint PicBack, PicShield, .x - 2, .y - 2, LEN_BLOCK, LEN_BLOCK, .ShieldID * LEN_BLOCK, 0, vbRed
                    .ShieldID = (.ShieldID + 1) Mod 2
                End If
            End With
        End If
    Next i
        
    For i = 0 To 25
        For j = 0 To 25
            If Map(i, j) = 2 Then
                Tomin i * WG, j * WG, WG, 0, 0, 0, WG, WG, PicBack.hdc, PicDX(1).hdc
            End If
        Next j
    Next i
    
    For i = 0 To 4
        With BornMovie(i)
            If .Show Then TransparentPaint PicBack, PicStar, .x, .y, LEN_BLOCK, LEN_BLOCK, (.No - 2) * LEN_BLOCK, 0, vbRed
        End With
    Next i
    
    For i = 0 To 19
        If Explode(i).Show Then
            With Explode(i)
                Tomin .x, .y, 2 * LEN_BLOCK, .No * 2 * LEN_BLOCK, 0, .No * 2 * LEN_BLOCK, 2 * LEN_BLOCK, 2 * LEN_BLOCK, PicBack.hdc, PicBZ(3).hdc
            End With
        End If
    Next i
    
    If Bonus.Show Then BitBlt PicBack.hdc, Bonus.x, Bonus.y, Bonus.W, Bonus.H, PicBonus(Bonus.T).hdc, 0, 0, vbSrcCopy
    
    If GameState = GS_GAMEOVER Then TransparentPaint PicBack, PicGameOver.Picture, 90, 130

    lblLife(0).Caption = PLLife(0)
    lblLife(1).Caption = PLLife(1)
End Sub


Public Sub GameOver(Optional ByVal ks As Boolean)
    If Not ks Then sndPlaySound "BANG.wav", 1
    GameState = GS_GAMEOVER
    Prevec = False
    If ks = 0 Then BitBlt PicBack.hdc, 12 * WG, 24 * WG, LEN_BLOCK, LEN_BLOCK, PicIcon(0).hdc, 0, 0, SRCCOPY
    TransparentPaint PicBack, PicGameOver.Picture, 90, 130
End Sub

Public Sub PaintETank(ByVal item As Integer)
    With EYTank(item)
        Select Case .TYPE
        Case 0
            TransparentPaint PicBack, Picetank.Picture, .x, .y, LEN_TANK, LEN_TANK, (IIf(EYMove(item), 1, 0) + IIf(.Move, 1, 0)) * LEN_TANK, .FANG * LEN_TANK, 0
        Case 1
            TransparentPaint PicBack, Picetank.Picture, .x, .y, LEN_TANK, LEN_TANK, (IIf(EYMove(item), 3, 2) + IIf(.Move, 1, 0)) * LEN_TANK, .FANG * LEN_TANK, 0
        Case 2
            TransparentPaint PicBack, Picetank.Picture, .x, .y, LEN_TANK, LEN_TANK, (IIf(EYMove(item), 5, 4) + IIf(.Move, 1, 0)) * LEN_TANK, .FANG * LEN_TANK, 0
        Case 3
            TransparentPaint PicBack, Picetank.Picture, .x, .y, LEN_TANK, LEN_TANK, (IIf(EYMove(item), 7, 6) + IIf(.Move, 1, 0)) * LEN_TANK, .FANG * LEN_TANK, 0
        Case 4
            TransparentPaint PicBack, Picetank.Picture, .x, .y, LEN_TANK, LEN_TANK, (((4 - .FHP) + IIf(EYMove(item), 1, 0)) * 2 + IIf(.Move, 1, 0)) * LEN_TANK, (.FANG + 4) * LEN_TANK, 0
        Case 5
            TransparentPaint PicBack, Picetank.Picture, .x, .y, LEN_TANK, LEN_TANK, ((IIf(4 - .FHP, 4 - .FHP, 4) + IIf(EYMove(item), 1, 0)) * 2 + IIf(.Move, 1, 0)) * LEN_TANK, (.FANG + 4) * LEN_TANK, 0
        Case 6
            TransparentPaint PicBack, Picetank.Picture, .x, .y, LEN_TANK, LEN_TANK, (IIf(EYMove(item), 9, 8) + IIf(.Move, 1, 0)) * LEN_TANK, .FANG * LEN_TANK, 0
        End Select
        If .Shield Then
            TransparentPaint PicBack, PicShield, .x - 2, .y - 2, LEN_BLOCK, LEN_BLOCK, .ShieldID * LEN_BLOCK, 0, vbRed
            .ShieldID = (.ShieldID + 1) Mod 2
        End If
    End With
End Sub

Private Sub RndEnemy_Timer()
    If GameState <> GS_GAMING And GameState <> GS_GAMEOVER Then Exit Sub
    Randomize Now 'Timer
    For i = 0 To 3
        If EYTank(i).Show Then
            With EYTank(i)
                .Move = Not .Move
                If (Rnd * 100) Mod 90 = 0 Then
                    .FANG = (.FANG + Int(100 * Rnd(i)) Mod 3) Mod 4
                End If
                If Int(400 * Rnd) <= (Setting.DEPTH ^ 2) * 2 And EBullet(MinEBullet).Show = False Then
                    EBullet(MinEBullet).FANG = .FANG
                    EBullet(MinEBullet).Show = True
                    EBullet(MinEBullet).x = .x + 10
                    EBullet(MinEBullet).y = .y + 10
                    If .TYPE > 5 Then
                        EBullet(MinEBullet).TYPE = 1
                        EBullet(MinEBullet).VD = 6
                    ElseIf .TYPE > 3 Then
                        EBullet(MinEBullet).TYPE = 0
                        EBullet(MinEBullet).VD = 6
                    ElseIf .TYPE > 1 Then
                        EBullet(MinEBullet).TYPE = 0
                        EBullet(MinEBullet).VD = 9
                    Else
                        EBullet(MinEBullet).TYPE = 0
                        EBullet(MinEBullet).VD = 8
                    End If
                    MinEBullet = (MinEBullet + 1) Mod 20
                End If
'                Debug.Assert (.TYPE <> 2 And .TYPE <> 3)
                Select Case .FANG
                Case TS_DOWN
                    xi = .x \ WG: xb = .x - xi * WG
                    yi = (.y + LEN_TANK) \ WG: yb = .y + LEN_TANK - yi * WG
                    If yb < .FER And yi < 26 Then
                        If Map(xi, yi) = 1 Or Map(xi, yi) = 3 Or Map(xi, yi) = 4 Or Map(xi, yi) = MAPZN + TS_RIGHT Or Map(xi, yi) = MAPZN + TS_UP Then GoTo IFCAN_T
                        If Map(xi, yi) = MAPZN + TS_LEFT And xb < HWG Then GoTo IFCAN_T
                        
                        If Map(xi + 1, yi) = 1 Or Map(xi + 1, yi) = 3 Or Map(xi + 1, yi) = 4 Or _
                            Map(xi + 1, yi) = MAPZN + TS_RIGHT Or Map(xi + 1, yi) = MAPZN + TS_LEFT _
                            Or Map(xi + 1, yi) = MAPZN + TS_UP Then GoTo IFCAN_T
                                                    
                        If xi < 24 Then If xb > 4 And (Map(xi + 2, yi) = 1 Or Map(xi + 2, yi) = 3 Or Map(xi + 2, yi) = 4 Or Map(xi + 2, yi) = MAPZN + TS_UP Or Map(xi + 2, yi) = MAPZN + TS_LEFT) Then GoTo IFCAN_T
                    End If
                    If yb >= HWG - .FER And yi < 24 Then
                        If Map(xi, yi) = MAPZN + TS_DOWN Then GoTo IFCAN_T
                        If Map(xi + 1, yi) = MAPZN + TS_DOWN Then GoTo IFCAN_T
                        If xi < 24 Then If xb > HWG And Map(xi + 2, yi) = MAPZN + TS_DOWN Then GoTo IFCAN_T
                    End If
                    For j = 0 To 3
                        If EYTank(j).Show And i <> j Then
                            If EYTank(j).y - .y < .FER + LEN_TANK And .y < EYTank(j).y _
                                And .x > EYTank(j).x - LEN_TANK And .x < EYTank(j).x + LEN_TANK Then GoTo IFCAN_T
                        End If
                    Next j
                    For j = 0 To 1
                        If PLTank(j).Show Then
                            If PLTank(j).y - .y < .FER + LEN_TANK And .y < PLTank(j).y _
                                And .x > PLTank(j).x - LEN_TANK And .x < PLTank(j).x + LEN_TANK Then GoTo IFCAN_T
                        End If
                    Next j
                    If .y < 420 - LEN_TANK - .FER Then .y = .y + .FER Else GoTo IFCAN_T
                Case TS_UP
                    xi = .x \ WG: xb = .x - xi * WG
                    yi = .y \ WG: yb = .y - yi * WG
                    If yb < .FER And yi > 0 Then
                        If Map(xi, yi - 1) = 1 Or Map(xi, yi - 1) = 3 Or Map(xi, yi - 1) = 4 Or Map(xi, yi - 1) = MAPZN + TS_RIGHT Or Map(xi, yi - 1) = MAPZN + TS_DOWN Then GoTo IFCAN_T
                        If Map(xi, yi - 1) = MAPZN + TS_LEFT And xb < HWG Then GoTo IFCAN_T
                        
                        If Map(xi + 1, yi - 1) = 1 Or Map(xi + 1, yi - 1) = 3 Or Map(xi + 1, yi - 1) = 4 Or _
                            Map(xi + 1, yi - 1) = MAPZN + TS_RIGHT Or Map(xi + 1, yi - 1) = MAPZN + TS_LEFT _
                            Or Map(xi + 1, yi - 1) = MAPZN + TS_DOWN Then GoTo IFCAN_T
                        
                        If xi < 24 Then If xb > 4 And (Map(xi + 2, yi - 1) = 1 Or Map(xi + 2, yi - 1) = 3 Or Map(xi + 2, yi - 1) = 4 Or Map(xi + 2, yi - 1) = MAPZN + TS_DOWN Or Map(xi + 2, yi - 1) = MAPZN + TS_LEFT) Then GoTo IFCAN_T
                    End If
                    If yb < HWG + .FER And yi > 0 Then
                        If Map(xi, yi) = MAPZN + TS_UP Then GoTo IFCAN_T
                        If Map(xi + 1, yi) = MAPZN + TS_UP Then GoTo IFCAN_T
                        If xi < 24 Then If xb > HWG And Map(xi + 2, yi) = MAPZN + TS_UP Then GoTo IFCAN_T
                    End If
                    
                    For j = 0 To 3
                        If EYTank(j).Show And i <> j Then
                            If .y - EYTank(j).y < .FER + LEN_TANK And .y > EYTank(j).y _
                                And .x > EYTank(j).x - LEN_TANK And .x < EYTank(j).x + LEN_TANK Then GoTo IFCAN_T
                        End If
                    Next j
                    For j = 0 To 1
                        If PLTank(j).Show Then
                            If .y - PLTank(j).y < .FER + LEN_TANK And .y > PLTank(j).y _
                                And .x > PLTank(j).x - LEN_TANK And .x < PLTank(j).x + LEN_TANK Then GoTo IFCAN_T
                        End If
                    Next j
                    If .y > .FER Then .y = .y - .FER Else GoTo IFCAN_T
                Case TS_LEFT
                    xi = .x \ WG: xb = .x - xi * WG
                    yi = .y \ WG: yb = .y - yi * WG
                    If xb < .FER And xi > 0 Then
                        If Map(xi - 1, yi) = 1 Or Map(xi - 1, yi) = 3 Or Map(xi - 1, yi) = 4 Or Map(xi - 1, yi) = MAPZN + TS_RIGHT Or Map(xi - 1, yi) = MAPZN + TS_DOWN Then GoTo IFCAN_T
                        If Map(xi - 1, yi) = MAPZN + TS_UP And yb < HWG Then GoTo IFCAN_T
                        
                        If Map(xi - 1, yi + 1) = 1 Or Map(xi - 1, yi + 1) = 3 Or Map(xi - 1, yi + 1) = 4 Or _
                            Map(xi - 1, yi + 1) = MAPZN + TS_RIGHT Or Map(xi - 1, yi + 1) = MAPZN + TS_UP _
                            Or Map(xi - 1, yi + 1) = MAPZN + TS_DOWN Then GoTo IFCAN_T
                        
                        If yi < 24 Then If yb > 4 And (Map(xi - 1, yi + 2) = 1 Or Map(xi - 1, yi + 2) = 3 Or Map(xi - 1, yi + 2) = 4 Or Map(xi - 1, yi + 2) = MAPZN + TS_UP Or Map(xi - 1, yi + 2) = MAPZN + TS_RIGHT) Then GoTo IFCAN_T
                    End If
                    If xb < HWG + .FER And xi > 0 Then
                        If Map(xi, yi) = MAPZN + TS_LEFT Then GoTo IFCAN_T
                        If Map(xi, yi + 1) = MAPZN + TS_LEFT Then GoTo IFCAN_T
                        If yi < 24 Then If yb > HWG And Map(xi, yi + 2) = MAPZN + TS_LEFT Then GoTo IFCAN_T
                    End If
                    For j = 0 To 3
                        If EYTank(j).Show Then
                            If EYTank(j).y <= .y + LEN_TANK And .y <= EYTank(j).y + LEN_TANK _
                                And .x - EYTank(j).x < .FER + LEN_TANK And .x > EYTank(j).x Then GoTo IFCAN_T
                        End If
                    Next j
                    For j = 0 To 1
                        If PLTank(j).Show Then
                            If PLTank(j).y <= .y + LEN_TANK And .y <= PLTank(j).y + LEN_TANK _
                                And .x - PLTank(j).x < .FER + LEN_TANK And .x > PLTank(j).x Then GoTo IFCAN_T
                        End If
                    Next j
                    If .x > .FER Then .x = .x - .FER Else GoTo IFCAN_T
                Case TS_RIGHT
                    xi = (.x + LEN_TANK) \ WG: xb = .x + LEN_TANK - xi * WG
                    yi = .y \ WG: yb = .y - yi * WG
                    If xb < .FER And xi < 26 Then
                        If Map(xi, yi) = 1 Or Map(xi, yi) = 3 Or Map(xi, yi) = 4 Or Map(xi, yi) = MAPZN + TS_LEFT Or Map(xi - 1, yi) = MAPZN + TS_DOWN Then GoTo IFCAN_T
                        If Map(xi, yi) = MAPZN + TS_UP And yb < HWG Then GoTo IFCAN_T
                        
                        If Map(xi, yi + 1) = 1 Or Map(xi, yi + 1) = 3 Or Map(xi, yi + 1) = 4 Or _
                            Map(xi, yi + 1) = MAPZN + TS_LEFT Or Map(xi, yi + 1) = MAPZN + TS_UP _
                            Or Map(xi, yi + 1) = MAPZN + TS_DOWN Then GoTo IFCAN_T
                        
                        If yi < 24 Then If yb > 4 And (Map(xi, yi + 2) = 1 Or Map(xi, yi + 2) = 3 Or Map(xi, yi + 2) = 4 Or Map(xi, yi + 2) = MAPZN + TS_UP Or Map(xi, yi + 2) = MAPZN + TS_LEFT) Then GoTo IFCAN_T
                    End If
                    If xb >= HWG + .FER And xi < 24 Then
                        If Map(xi, yi) = MAPZN + TS_RIGHT Then GoTo IFCAN_T
                        If Map(xi, yi + 1) = MAPZN + TS_RIGHT Then GoTo IFCAN_T
                        If yi < 24 Then If yb > HWG And Map(xi, yi + 2) = MAPZN + TS_RIGHT Then GoTo IFCAN_T
                    End If
                    For j = 0 To 3
                        If EYTank(j).Show Then
                            If EYTank(j).y <= .y + LEN_TANK And .y <= EYTank(j).y + LEN_TANK _
                                And EYTank(j).x - .x < .FER + LEN_TANK And .x < EYTank(j).x Then GoTo IFCAN_T
                        End If
                    Next j
                    For j = 0 To 1
                        If PLTank(j).Show Then
                            If PLTank(j).y <= .y + LEN_TANK And .y <= PLTank(j).y + LEN_TANK _
                                And PLTank(j).x - .x < .FER + LEN_TANK And .x < PLTank(j).x Then GoTo IFCAN_T
                        End If
                    Next j
                    If .x < 420 - LEN_TANK - .FER Then .x = .x + .FER Else GoTo IFCAN_T
                End Select
                If Bonus.Show And (.TYPE Mod 2 = 1) Then
                    If Bonus.y <= .y + LEN_BLOCK And .y <= Bonus.y + LEN_BLOCK Then
                        If .x >= Bonus.x - LEN_BLOCK And .x <= Bonus.x + LEN_BLOCK Then
                            Bonus.Show = False
                            EchoBonus Bonus.T, 2 + i
                        End If
                    End If
                End If
                GoTo ENDIFCAN_T
IFCAN_T:
'                If EBullet(MinEBullet).Show = True Then GoTo ENDIFCAN_T
                EBullet(MinEBullet).FANG = .FANG
                EBullet(MinEBullet).Show = True
                'If .FANG = TS_UP Or .FANG = TS_DOWN Then Else EBullet(MinEBullet).x = .xEBullet(MinEBullet).TYPE = IIf(.TYPE > 5, 1, 0)If .FANG = TS_LEFT Or .FANG = TS_RIGHT ThenElse EBullet(MinEBullet).y = .y
                EBullet(MinEBullet).x = .x + 10
                EBullet(MinEBullet).y = .y + 10
                If .TYPE > 5 Then
                    EBullet(MinEBullet).TYPE = 1
                    EBullet(MinEBullet).VD = 6
                ElseIf .TYPE > 3 Then
                    EBullet(MinEBullet).TYPE = 0
                    EBullet(MinEBullet).VD = 6
                ElseIf .TYPE > 1 Then
                    EBullet(MinEBullet).TYPE = 0
                    EBullet(MinEBullet).VD = 9
                Else
                    EBullet(MinEBullet).TYPE = 0
                    EBullet(MinEBullet).VD = 8
                End If
                MinEBullet = (MinEBullet + 1) Mod 20
                'tmpfang = .FANG
               ' While tmpfang = .FANG
                    .FANG = (99 * Rnd(.FANG + 1)) Mod 4 'Int(100 * Rnd(i)) Mod 3
                'Wend

ENDIFCAN_T:
            End With
        End If
    Next i
End Sub

Public Function FRnd(ByVal Index As Integer) As Integer
    Randomize Timer
    Static Stack(3, 3) As Integer
    Static StackTop(3) As Integer
    Dim i As Integer, max As Integer, min As Integer, ni As Integer, mi As Integer
    Dim tmp As Integer
    tmp = Int(3 * Rnd)
    For i = 0 To 3
        If min > Stack(Index, i) Then min = Stack(Index, i): ni = i
        If max < Stack(Index, i) Then max = Stack(Index, i): mi = i
    Next i
    If max - min > 4 \ Int(Sqr(StackTop(Index)) + 1) Then
        FRnd = ni
        Stack(Index, ni) = Stack(Index, ni) + 1
        StackTop(ni) = StackTop(ni) + 1
    ElseIf max - min > 3 \ Int(Sqr(StackTop(Index)) + 1) Then
        If tmp = ni Then
            FRnd = tmp
            Stack(Index, tmp) = Stack(Index, tmp) + 1
            StackTop(tmp) = StackTop(tmp) + 1
        Else
            tmp = Int(3 * Rnd)
            FRnd = tmp
            Stack(Index, tmp) = Stack(Index, tmp) + 1
            StackTop(tmp) = StackTop(tmp) + 1
        End If
    Else
        FRnd = tmp
        Stack(Index, tmp) = Stack(Index, tmp) + 1
        StackTop(tmp) = StackTop(tmp) + 1
    End If
End Function

Private Sub SaveM_Click()
    CoDi.Filter = "游戏进度文件(.gpf)|*.gpf"
    CoDi.ShowSave
    If CoDi.FileName <> "" Then
        Open CoDi.FileName For Binary Access Write As 1
            Put #1, , Map
            Put #1, , PLTank
            Put #1, , PBullet
            Put #1, , EYTank
            Put #1, , EBullet
            Put #1, , Bonus
            Put #1, , Remain
            Put #1, , NumPresent
            Put #1, , GameState
            Put #1, , PLLife
        Close #1
    End If
End Sub

Private Sub MnuSetting_Click()
    frmSet.Show
End Sub

Private Sub SaveP_Click()
    Dim cDib As New cDIBSection
    CoDi.Filter = "BMP位图|*.bmp|JEPG压缩位图|*.jpg"
    CoDi.InitDir = App_Path + "Image\"
    CoDi.ShowSave
    If CoDi.FileName <> "" Then
        If CoDi.FilterIndex = 1 Then
            SavePicture PicBack.Image, CoDi.FileName + ".bmp"
        Else
            cDib.CreateFromPicture PicBack.Picture
            SaveJPG cDib, CoDi.FileName + ".jpg"
        End If
    End If
End Sub

Private Sub SetupTo_Click()
    On Error Resume Next
    Dim Setting As Configure
    Dim APP2() As Byte
    Dim Counter As Long
        
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
        .DEPTH = 0
    End With
    Open "Tank.cfg" For Binary Access Write As 1 Len = Len(Setting)
    Put #1, , Setting
    Close
    
    ReleaseMap

    APP2 = LoadResData("BANG", "WAVE")
    Counter = UBound(APP2)
    ReDim Preserve APP2(Counter) As Byte
    Open "bang.wav " For Binary Access Write As #1
        Put #1, , APP2
    Close #1
    
    APP2 = LoadResData("FARE", "WAVE")
    Counter = UBound(APP2)
    ReDim Preserve APP2(Counter) As Byte
    Open "Fanfare.wav " For Binary Access Write As #1
        Put #1, , APP2
    Close #1
    
    APP2 = LoadResData("FIRE", "WAVE")
    Counter = UBound(APP2)
    ReDim Preserve APP2(Counter) As Byte
    Open "fire.wav " For Binary Access Write As #1
        Put #1, , APP2
    Close #1
    
    APP2 = LoadResData("hit", "WAVE")
    Counter = UBound(APP2)
    ReDim Preserve APP2(Counter) As Byte
    Open "hit.wav " For Binary Access Write As #1
        Put #1, , APP2
    Close #1
    
    APP2 = LoadResData("peow", "WAVE")
    Counter = UBound(APP2)
    ReDim Preserve APP2(Counter) As Byte
    Open "peow.wav " For Binary Access Write As #1
        Put #1, , APP2
    Close #1
    
    
    MkDir "Image"
    MkDir "Temp"
End Sub

Private Sub SleepBonus_Timer()
    If GameState <> GS_GAMING Then Exit Sub
    For i = 0 To 2
        If EatBonus(i) Then
            If i = 0 Then
                Map(11, 25) = ZHUANQIANG
                Map(11, 24) = ZHUANQIANG
                Map(11, 23) = ZHUANQIANG
                Map(12, 23) = ZHUANQIANG
                Map(13, 23) = ZHUANQIANG
                Map(14, 24) = ZHUANQIANG
                Map(14, 25) = ZHUANQIANG
                Map(14, 23) = ZHUANQIANG
                EatBonus(0) = 0
            ElseIf i = 1 Then
                RndEnemy.Enabled = True
                Me.KeyPreview = True
                EatBonus(1) = 0
            ElseIf i = 2 Then
                If EatBonus(2) <= 2 Then
                    PLTank(EatBonus(2) - 1).Shield = False
                Else
                    EYTank(EatBonus(2) - 3).Shield = False
                End If
                EatBonus(2) = 0
            End If
        End If
    Next i
    SleepBonus.Enabled = False
End Sub

Private Sub Sleeper_Timer()
    Select Case SleepType
    Case GE_WIN
        If GameState <> GS_GAMING And GameState <> GS_PAUSE Then Sleeper.Enabled = False
        PicBack.Picture = LoadPicture()
        GameState = GS_RESULT
        ShowRemark
        SleepType = GE_START
        Sleeper.Interval = 3000
        Prevec = True
        Exit Sub
    Case GE_START
        n = FileLen("map.bin") \ SectionLen
        If Level = n Then Level = 0
        Start Level + 1
    End Select
    Sleeper.Enabled = False
End Sub

Private Sub Stop_Click()
    If GameState = GS_PAUSE Then GameState = GS_GAMING: Exit Sub
    If GameState = GS_GAMING Then GameState = GS_PAUSE
End Sub

Private Sub Update_Timer()
    If GameState <> GS_GAMING And GameState <> GS_GAMEOVER Then Exit Sub
    PicBack.Picture = LoadPicture()
    PaintScreen
    For i = 0 To 19
        If Explode(i).Show Then
            With Explode(i)
                If .No = .NUM Then .Show = False
                .No = .No + 1
            End With
        End If
    Next i
    
    For Tndex = 0 To 1
        If PLTank(Tndex).Show Then
            With PLTank(Tndex)
                xc = .x \ WG: yc = .y \ WG
                If Map(xc, yc) = BINGDI Or Map(xc, yc + 1) = BINGDI Or Map(xc + 1, yc) = BINGDI Or Map(xc + 1, yc + 1) = BINGDI Then
                    If .FANG = TS_DOWN Then
                        xi = xc: xb = .x - xi * WG
                        yi = (.y + LEN_TANK) \ WG: yb = .y + LEN_TANK - yi * WG
                        If yb = 0 And yi < 26 Then
                            If Map(xi, yi) = 1 Or Map(xi, yi) = 3 Or Map(xi, yi) = 4 Then GoTo LT
                            If Map(xi, yi) = MAPZN + TS_LEFT And xb < HWG Then GoTo LT
                            If Map(xi, yi) = MAPZN + TS_RIGHT Or Map(xi, yi) = MAPZN + TS_UP Then GoTo LT

                            If Map(xi + 1, yi) = 1 Or Map(xi + 1, yi) = 3 Or Map(xi + 1, yi) = 4 Or _
                                Map(xi + 1, yi) = MAPZN + TS_RIGHT Or Map(xi + 1, yi) = MAPZN + TS_LEFT _
                                Or Map(xi + 1, yi) = MAPZN + TS_UP Then GoTo LT
                                                        
                            If xi < 24 Then If xb > 4 And (Map(xi + 2, yi) = 1 Or Map(xi + 2, yi) = 3 Or Map(xi + 2, yi) = 4 Or Map(xi + 2, yi) = MAPZN + TS_UP Or Map(xi + 2, yi) = MAPZN + TS_LEFT) Then GoTo LT
                        End If
                        If yb = HWG And yi < 26 Then
                            If Map(xi, yi) = MAPZN + TS_DOWN Then GoTo LT
                            If Map(xi + 1, yi) = MAPZN + TS_DOWN Then GoTo LT
                            If xi < 24 Then If xb > HWG And Map(xi + 2, yi) = MAPZN + TS_DOWN Then GoTo LT
                        End If
                        For i = 0 To 3
                            If EYTank(i).Show Then
                                If EYTank(i).y - .y < 4 + LEN_TANK And .y < EYTank(i).y _
                                    And .x > EYTank(i).x - LEN_TANK And .x < EYTank(i).x + LEN_TANK Then GoTo LT
                            End If
                        Next i
                        If PLTank(1 - Tndex).Show Then
                            If PLTank(1 - Tndex).y - .y < 4 + LEN_TANK And .y < PLTank(1 - Tndex).y _
                                And .x > PLTank(1 - Tndex).x - LEN_TANK And .x < PLTank(1 - Tndex).x + LEN_TANK Then GoTo LT
                        End If
                        If .y < 388 Then .y = .y + 4
                        .Move = True
                    ElseIf .FANG = TS_UP Then
                        xi = xc: xb = .x - xi * WG
                        yi = yc: yb = .y - yi * WG
                        If yb = 0 And yi > 0 Then
                            If Map(xi, yi - 1) = 1 Or Map(xi, yi - 1) = 3 Or Map(xi, yi - 1) = 4 Then GoTo LT
                            If Map(xi, yi - 1) = MAPZN + TS_LEFT And xb < HWG Then GoTo LT
                            If Map(xi, yi - 1) = MAPZN + TS_RIGHT Or Map(xi, yi - 1) = MAPZN + TS_DOWN Then GoTo LT
                            
                            If Map(xi + 1, yi - 1) = 1 Or Map(xi + 1, yi - 1) = 3 Or Map(xi + 1, yi - 1) = 4 Or _
                                Map(xi + 1, yi - 1) = MAPZN + TS_RIGHT Or Map(xi + 1, yi - 1) = MAPZN + TS_LEFT _
                                Or Map(xi + 1, yi - 1) = MAPZN + TS_DOWN Then GoTo LT
                            
                            If xi < 24 Then If xb > 4 And (Map(xi + 2, yi - 1) = 1 Or Map(xi + 2, yi - 1) = 3 Or Map(xi + 2, yi - 1) = 4 Or Map(xi + 2, yi - 1) = MAPZN + TS_DOWN Or Map(xi + 2, yi - 1) = MAPZN + TS_LEFT) Then GoTo LT
                        End If
                        If yb = HWG And yi > 0 Then
                            If Map(xi, yi) = MAPZN + TS_UP Then GoTo LT
                            If Map(xi + 1, yi) = MAPZN + TS_UP Then GoTo LT
                            If xi < 24 Then If xb > HWG And Map(xi + 2, yi) = MAPZN + TS_UP Then GoTo LT
                        End If
                        For i = 0 To 3
                            If EYTank(i).Show Then
                                If .y - EYTank(i).y < 4 + LEN_TANK And .y > EYTank(i).y _
                                    And .x > EYTank(i).x - LEN_TANK And .x < EYTank(i).x + LEN_TANK Then GoTo LT
                            End If
                        Next i
                        If PLTank(1 - Tndex).Show Then
                            If .y - PLTank(1 - Tndex).y < 4 + LEN_TANK And .y > PLTank(1 - Tndex).y _
                                And .x > PLTank(1 - Tndex).x - LEN_TANK And .x < PLTank(1 - Tndex).x + LEN_TANK Then GoTo LT
                        End If
                        If .y > 3 Then .y = .y - 4
                        .Move = True
                    ElseIf .FANG = TS_LEFT Then
                        xi = xc: xb = .x - xi * WG
                        yi = yc: yb = .y - yi * WG
                        If xb = 0 And xi > 0 Then
                            If Map(xi - 1, yi) = 1 Or Map(xi - 1, yi) = 3 Or Map(xi - 1, yi) = 4 Then GoTo LT
                            If Map(xi - 1, yi) = MAPZN + TS_UP And yb < HWG Then GoTo LT
                            If Map(xi - 1, yi) = MAPZN + TS_RIGHT Or Map(xi - 1, yi) = MAPZN + TS_DOWN Then GoTo LT
                            
                            If Map(xi - 1, yi + 1) = 1 Or Map(xi - 1, yi + 1) = 3 Or Map(xi - 1, yi + 1) = 4 Or _
                                Map(xi - 1, yi + 1) = MAPZN + TS_RIGHT Or Map(xi - 1, yi + 1) = MAPZN + TS_UP _
                                Or Map(xi - 1, yi + 1) = MAPZN + TS_DOWN Then GoTo LT
                            
                            If yi < 24 Then If yb > 4 And (Map(xi - 1, yi + 2) = 1 Or Map(xi - 1, yi + 2) = 3 Or Map(xi - 1, yi + 2) = 4 Or Map(xi - 1, yi + 2) = MAPZN + TS_UP Or Map(xi - 1, yi + 2) = MAPZN + TS_RIGHT) Then GoTo LT
                        End If
                        If xb = HWG And xi > 0 Then
                            If Map(xi, yi) = MAPZN + TS_LEFT Then GoTo LT
                            If Map(xi, yi + 1) = MAPZN + TS_LEFT Then GoTo LT
                            If yi < 24 Then If yb > HWG And Map(xi, yi + 2) = MAPZN + TS_LEFT Then GoTo LT
                        End If
                        For i = 0 To 3
                            If EYTank(i).Show Then
                                If EYTank(i).y <= .y + LEN_TANK And .y <= EYTank(i).y + LEN_TANK _
                                    And .x - EYTank(i).x < 4 + LEN_TANK And .x > EYTank(i).x Then GoTo LT
                            End If
                        Next i
                        If PLTank(1 - Tndex).Show Then
                            If PLTank(1 - Tndex).y <= .y + LEN_TANK And .y <= PLTank(1 - Tndex).y + LEN_TANK _
                                And .x - PLTank(1 - Tndex).x < 4 + LEN_TANK And .x > PLTank(1 - Tndex).x Then GoTo LT
                        End If
                        If .x > 3 Then .x = .x - 4
                        .Move = True
                    ElseIf .FANG = TS_RIGHT Then
                        xi = (.x + LEN_TANK) \ WG: xb = .x + LEN_TANK - xi * WG
                        yi = .y \ WG: yb = .y - yi * WG
                        If xb = 0 And xi < 26 Then
                            If Map(xi, yi) = 1 Or Map(xi, yi) = 3 Or Map(xi, yi) = 4 Then GoTo LT
                            If Map(xi, yi) = MAPZN + TS_UP And yb < HWG Then GoTo LT
                            If Map(xi, yi) = MAPZN + TS_LEFT Or Map(xi - 1, yi) = MAPZN + TS_DOWN Then GoTo LT
                            
                            If Map(xi, yi + 1) = 1 Or Map(xi, yi + 1) = 3 Or Map(xi, yi + 1) = 4 Or _
                                Map(xi, yi + 1) = MAPZN + TS_LEFT Or Map(xi, yi + 1) = MAPZN + TS_UP _
                                Or Map(xi, yi + 1) = MAPZN + TS_DOWN Then GoTo LT
                            
                            If yi < 24 Then If yb > 4 And (Map(xi, yi + 2) = 1 Or Map(xi, yi + 2) = 3 Or Map(xi, yi + 2) = 4 Or Map(xi, yi + 2) = MAPZN + TS_UP Or Map(xi, yi + 2) = MAPZN + TS_LEFT) Then GoTo LT
                        End If
                        If xb = HWG And xi < 26 Then
                            If Map(xi, yi) = MAPZN + TS_RIGHT Then GoTo LT
                            If Map(xi, yi + 1) = MAPZN + TS_RIGHT Then GoTo LT
                            If yi < 24 Then If yb > HWG And Map(xi, yi + 2) = MAPZN + TS_RIGHT Then GoTo LT
                        End If
                        For i = 0 To 3
                            If EYTank(i).Show Then
                                If EYTank(i).y <= .y + LEN_TANK And .y <= EYTank(i).y + LEN_TANK _
                                    And EYTank(i).x - .x < 4 + LEN_TANK And .x < EYTank(i).x Then GoTo LT
                            End If
                        Next i
                        If PLTank(1 - Tndex).Show Then
                            If PLTank(1 - Tndex).y <= .y + LEN_TANK And .y <= PLTank(1 - Tndex).y + LEN_TANK _
                                And PLTank(1 - Tndex).x - .x < 4 + LEN_TANK And .x < PLTank(1 - Tndex).x Then GoTo LT
                        End If
                        If .x < 388 Then .x = .x + 4
                        .Move = True
                    End If
                End If
LT:         End With
        End If
    Next Tndex

End Sub

Public Sub Sleep(ByVal typ As Integer, ByVal sec As Integer)
    If typ < 5 Then
        Sleeper.Interval = 1000 * sec
        Sleeper.Enabled = True
        SleepType = typ
    Else
        SleepBonus.Interval = 1000 * sec
        SleepBonus.Enabled = True
    End If
End Sub

Public Sub EchoBonus(ByVal typ As Integer, ByVal plr As Integer)
    Dim a As Byte
    Select Case typ
    Case 0
        If plr < 2 Then
            Map(11, 25) = TIEQIANG
            Map(11, 24) = TIEQIANG
            Map(11, 23) = TIEQIANG
            Map(12, 23) = TIEQIANG
            Map(13, 23) = TIEQIANG
            Map(14, 24) = TIEQIANG
            Map(14, 25) = TIEQIANG
            Map(14, 23) = TIEQIANG
            Sleep 8, 30
        Else
            Map(11, 25) = 0
            Map(11, 24) = 0
            Map(11, 23) = 0
            Map(12, 23) = 0
            Map(13, 23) = 0
            Map(14, 24) = 0
            Map(14, 25) = 0
            Map(14, 23) = 0
            Sleep 8, 20
        End If
        EatBonus(0) = True
        sndPlaySound "FANFARE.wav", 1
    Case 1
        If plr < 2 Then
            RndEnemy.Enabled = False
            Sleep 9, 15
        Else
            Me.KeyPreview = False
            Sleep 9, 8
        End If
        EatBonus(1) = True
    Case 2
        If plr < 2 Then
            PLTank(plr).Shield = True
        Else
            EYTank(plr - 2).Shield = True
        End If
        
        Sleep 10, 15
        EatBonus(2) = plr + 1
        sndPlaySound "FANFARE.wav", 1
    Case 3          '对方全部爆炸的
        If plr < 2 Then
            For i = 0 To 3
                If EYTank(i).Show Then
                    EYTank(i).Show = False
                    Explode(MinExplode).Show = True
                    Explode(MinExplode).No = 0
                    Explode(MinExplode).NUM = 6
                    Explode(MinExplode).x = EYTank(i).x - HWG
                    Explode(MinExplode).y = EYTank(i).y - HWG
                    Remark(plr, EYTank(i).TYPE \ 2) = Remark(plr, EYTank(i).TYPE \ 2) + 1
                    MinExplode = (MinExplode + 1) Mod 20
                    NumPresent = NumPresent - 1
                End If
            Next i
                    sndPlaySound "BANG.wav", 1
            
            If NumPresent = 0 Then
                If Remain <= 0 Then
                    Sleep GE_WIN, 2
                End If
            ElseIf NumPresent < 4 Then
                If Remain Then
                    BornMovie(Int(100 * Rnd) Mod 3).No = 1
                    BornKeeper.Enabled = True
                End If
            End If
        Else
            For i = 0 To 1
                If PLTank(i).Show Then
                    PLTank(i).Show = False
                    Explode(MinExplode).Show = True
                    Explode(MinExplode).No = 0
                    Explode(MinExplode).NUM = 6
                    Explode(MinExplode).x = PLTank(i).x - HWG
                    Explode(MinExplode).y = PLTank(i).y - HWG
                    MinExplode = (MinExplode + 1) Mod 20
                    sndPlaySound "BANG.wav", 1
                    PLLife(i) = PLLife(0) - 1
                End If
            Next i
            
            If PlayerNum = 1 Then
                If PLLife(0) = 0 Then
                    GameOver False
                Else
                    BornMovie(3).No = 1
                    BornKeeper.Enabled = True
                End If
            Else
                If PLLife(0) = 0 Then
                    a = 1
                Else
                    BornMovie(3).No = 1
                    BornKeeper.Enabled = True
                End If
                If PLLife(1) = 0 Then
                    a = a + 1
                Else
                    BornMovie(4).No = 1
                    BornKeeper.Enabled = True
                End If
                If a = 2 Then GameOver False
            End If
        End If
        BornKeeper.Enabled = True

    Case 4      '白星
        If plr < 2 Then
            PLTank(plr).FHP = PLTank(plr).FHP + 1
            If PLTank(plr).FHP > 5 Then PLTank(plr).FHP = 5
            If PLTank(plr).FHP > 2 Then PLTank(plr).ProWa = True
        Else
            EYTank(plr - 2).FER = 10
        End If
        sndPlaySound "peow.wav", 1
    Case 5
        If plr < 2 Then
            PLLife(plr) = PLLife(plr) + 1
        Else
            If NumPresent < 4 And Remain = 0 Then
                BornMovie(Int(10 * Rnd) Mod 3).No = 1
                BornKeeper.Enabled = True
            Else
                Image3(Remain).Visible = True
                Remain = Remain + 1
            End If
        End If
        sndPlaySound "FANFARE.wav", 1
    Case 6  '金星
        If plr < 2 Then
            PLTank(plr).FHP = PLTank(plr).FHP + 2
            If PLTank(plr).FHP > 5 Then PLTank(plr).FHP = 5
            If PLTank(plr).FHP > 2 Then PLTank(plr).ProWa = True
        End If
        sndPlaySound "peow.wav", 1
    End Select
End Sub

Public Sub ShowRemark()
'b 72, 124, 182, 238, 292,210+78
    BitBlt PicBack.hdc, 0, 0, 420, 420, PicWin(0).hdc, 0, 0, vbSrcCopy
    
    If Remark(0, 0) > 9 Then BitBlt PicBack.hdc, 70, 124, 14, 14, PicNum.hdc, (Remark(0, 0) \ 10) * 14, 0, vbSrcCopy
    BitBlt PicBack.hdc, 84, 124, 14, 14, PicNum.hdc, (Remark(0, 0) Mod 10) * 14, 0, vbSrcCopy
    
    If Remark(0, 1) > 9 Then BitBlt PicBack.hdc, 70, 182, 14, 14, PicNum.hdc, (Remark(0, 1) \ 10) * 14, 0, vbSrcCopy
    BitBlt PicBack.hdc, 84, 182, 14, 14, PicNum.hdc, (Remark(0, 1) Mod 10) * 14, 0, vbSrcCopy
    
    If Remark(0, 3) > 9 Then BitBlt PicBack.hdc, 70, 238, 14, 14, PicNum.hdc, (Remark(0, 3) \ 10) * 14, 0, vbSrcCopy
    BitBlt PicBack.hdc, 84, 238, 14, 14, PicNum.hdc, (Remark(0, 3) Mod 10) * 14, 0, vbSrcCopy
    
    If Remark(0, 2) > 9 Then BitBlt PicBack.hdc, 70, 292, 14, 14, PicNum.hdc, (Remark(0, 2) \ 10) * 14, 0, vbSrcCopy
    BitBlt PicBack.hdc, 84, 292, 14, 14, PicNum.hdc, (Remark(0, 2) Mod 10) * 14, 0, vbSrcCopy
    
    If PlayerNum = 1 Then
        BitBlt PicBack.hdc, 212, 40, 420, 420, PicWin(1).hdc, 0, 0, vbSrcCopy
        
        If Remark(1, 0) > 9 Then BitBlt PicBack.hdc, 290, 124, 14, 14, PicNum.hdc, (Remark(1, 0) \ 10) * 14, 0, vbSrcCopy
        BitBlt PicBack.hdc, 306, 124, 14, 14, PicNum.hdc, (Remark(1, 0) Mod 10) * 14, 0, vbSrcCopy
        
        If Remark(1, 1) > 9 Then BitBlt PicBack.hdc, 290, 182, 14, 14, PicNum.hdc, (Remark(1, 1) \ 10) * 14, 0, vbSrcCopy
        BitBlt PicBack.hdc, 306, 182, 14, 14, PicNum.hdc, (Remark(1, 1) Mod 10) * 14, 0, vbSrcCopy
        
        If Remark(1, 3) > 9 Then BitBlt PicBack.hdc, 290, 238, 14, 14, PicNum.hdc, (Remark(1, 3) \ 10) * 14, 0, vbSrcCopy
        BitBlt PicBack.hdc, 306, 238, 14, 14, PicNum.hdc, (Remark(1, 3) Mod 10) * 14, 0, vbSrcCopy
        
        If Remark(1, 2) > 9 Then BitBlt PicBack.hdc, 290, 292, 14, 14, PicNum.hdc, (Remark(1, 2) \ 10) * 14, 0, vbSrcCopy
        BitBlt PicBack.hdc, 306, 292, 14, 14, PicNum.hdc, (Remark(1, 2) Mod 10) * 14, 0, vbSrcCopy
    End If
End Sub

Public Function RndType() As Integer
    Dim test As Integer
    Randomize Now
    test = Int(1000 * Rnd)
    Select Case test
    Case Is < 200
        RndType = 0
    Case Is < 330
        RndType = 1
    Case Is < 480
        RndType = 2
    Case Is < 590
        RndType = 3
    Case Is < 740
        RndType = 4
    Case Is < 900
        RndType = 5
    Case Else
        RndType = 6
    End Select
End Function

Public Sub RefreshSetting()
    Open "Tank.cfg" For Binary Access Read As 1 Len = Len(Setting)
        Get #1, , Setting
    Close #1
End Sub
