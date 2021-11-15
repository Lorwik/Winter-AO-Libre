VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.ocx"
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   8985
   ClientLeft      =   1260
   ClientTop       =   1725
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":0ECA
   ScaleHeight     =   601.003
   ScaleMode       =   0  'User
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   7800
      Top             =   2040
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   0   'False
      Backlog         =   1
      Binary          =   0   'False
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   2048
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   999999
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.PictureBox Minimap 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   6840
      Picture         =   "frmMain.frx":3E467
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   42
      Top             =   120
      Width           =   1500
      Begin VB.Shape UserArea 
         BorderColor     =   &H80000004&
         Height          =   225
         Left            =   600
         Top             =   600
         Width           =   300
      End
      Begin VB.Shape UserM 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   75
         Left            =   705
         Shape           =   3  'Circle
         Top             =   690
         Width           =   60
      End
      Begin VB.Shape GuerraPos 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         Height          =   255
         Left            =   600
         Shape           =   3  'Circle
         Top             =   600
         Width           =   255
      End
   End
   Begin VB.Timer SecureLwK 
      Interval        =   30000
      Left            =   7800
      Top             =   2520
   End
   Begin VB.PictureBox Macros 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   5
      Left            =   7575
      ScaleHeight     =   450
      ScaleWidth      =   495
      TabIndex        =   39
      Top             =   8400
      Width           =   495
   End
   Begin VB.PictureBox Macros 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   4
      Left            =   6945
      ScaleHeight     =   450
      ScaleWidth      =   495
      TabIndex        =   38
      Top             =   8400
      Width           =   495
   End
   Begin VB.PictureBox Macros 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   3
      Left            =   6315
      ScaleHeight     =   450
      ScaleWidth      =   495
      TabIndex        =   37
      Top             =   8400
      Width           =   495
   End
   Begin VB.PictureBox Macros 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   2
      Left            =   5670
      ScaleHeight     =   450
      ScaleWidth      =   495
      TabIndex        =   36
      Top             =   8400
      Width           =   495
   End
   Begin VB.PictureBox Macros 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   1
      Left            =   5025
      ScaleHeight     =   450
      ScaleWidth      =   495
      TabIndex        =   35
      Top             =   8400
      Width           =   495
   End
   Begin VB.PictureBox Clima 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   7230
      Picture         =   "frmMain.frx":458AD
      ScaleHeight     =   255
      ScaleWidth      =   1095
      TabIndex        =   32
      Top             =   1665
      Width           =   1095
      Begin VB.Label lblTemp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0º"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   165
         Left            =   795
         TabIndex        =   40
         Top             =   0
         Width           =   195
      End
      Begin VB.Label Dia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Dia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   165
         Left            =   390
         TabIndex        =   33
         Top             =   120
         Width           =   240
      End
   End
   Begin VB.PictureBox Label2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   6840
      Picture         =   "frmMain.frx":46D79
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   31
      Top             =   120
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Timer WorkMacro 
      Enabled         =   0   'False
      Interval        =   800
      Left            =   7320
      Top             =   2040
   End
   Begin VB.Timer ActualizadorPosicion 
      Enabled         =   0   'False
      Left            =   6840
      Top             =   2040
   End
   Begin VB.TextBox SendCMSTXT 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   285
      Left            =   165
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1665
      Visible         =   0   'False
      Width           =   6990
   End
   Begin VB.TextBox SendTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   285
      Left            =   165
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1665
      Visible         =   0   'False
      Width           =   6975
   End
   Begin VB.CommandButton DespInv 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   9000
      MouseIcon       =   "frmMain.frx":4E1BF
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   3480
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.CommandButton DespInv 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   9000
      MouseIcon       =   "frmMain.frx":4E311
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   5400
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.ListBox hlst 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2205
      Left            =   9000
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3480
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2400
      Left            =   9000
      Picture         =   "frmMain.frx":4E463
      ScaleHeight     =   156.098
      ScaleMode       =   0  'User
      ScaleWidth      =   154.217
      TabIndex        =   1
      Top             =   3420
      Width           =   2400
   End
   Begin VB.Timer Second 
      Enabled         =   0   'False
      Interval        =   1050
      Left            =   7320
      Top             =   2520
   End
   Begin RichTextLib.RichTextBox RecTxt 
      Height          =   1500
      Left            =   120
      TabIndex        =   41
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   120
      Width           =   6630
      _ExtentX        =   11695
      _ExtentY        =   2646
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      OLEDropMode     =   0
      TextRTF         =   $"frmMain.frx":610A7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image img_soporte 
      Height          =   330
      Left            =   11040
      Picture         =   "frmMain.frx":61124
      Top             =   523
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Label sGM 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TIENES UNA RESPUESTA DE $NickGM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   8520
      MouseIcon       =   "frmMain.frx":6173E
      TabIndex        =   44
      Top             =   6360
      Visible         =   0   'False
      Width           =   3300
   End
   Begin VB.Label LvlLbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8805
      TabIndex        =   43
      Top             =   1620
      Width           =   420
   End
   Begin VB.Image Image7 
      Height          =   480
      Left            =   11400
      Picture         =   "frmMain.frx":62408
      Top             =   360
      Width           =   480
   End
   Begin VB.Image combateII 
      Height          =   255
      Left            =   8505
      Picture         =   "frmMain.frx":62CD2
      ToolTipText     =   "Modo Combate"
      Top             =   568
      Width           =   300
   End
   Begin VB.Image combate 
      Height          =   255
      Left            =   8505
      Picture         =   "frmMain.frx":63110
      ToolTipText     =   "Modo Combate"
      Top             =   568
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image Image6 
      Height          =   255
      Left            =   10320
      Top             =   8160
      Width           =   1335
   End
   Begin VB.Label online 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   10200
      TabIndex        =   34
      Top             =   630
      Width           =   105
   End
   Begin VB.Label GldLbl2 
      BackStyle       =   0  'Transparent
      Caption         =   "0 kk"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   10560
      TabIndex        =   30
      Top             =   6720
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Image PicSeg 
      Height          =   255
      Left            =   8820
      Picture         =   "frmMain.frx":6354E
      ToolTipText     =   "Seguro"
      Top             =   554
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label Agilidad 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   11445
      TabIndex        =   29
      Top             =   8520
      Width           =   105
   End
   Begin VB.Label Fuerza 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   10680
      TabIndex        =   28
      Top             =   8520
      Width           =   105
   End
   Begin VB.Label ItemName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "[None]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9000
      TabIndex        =   27
      Top             =   5880
      Width           =   2415
   End
   Begin VB.Image Image5 
      Height          =   135
      Left            =   0
      Top             =   0
      Width           =   12015
   End
   Begin VB.Label lblPorcLvl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   165
      Left            =   10410
      TabIndex        =   26
      Top             =   1965
      Width           =   525
   End
   Begin VB.Label exp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   165
      Left            =   10455
      TabIndex        =   25
      Top             =   1965
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Shape ExpShp 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   9765
      Top             =   1980
      Width           =   1785
   End
   Begin VB.Label lblMapaName 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Mapa desconocido"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   9360
      TabIndex        =   24
      Top             =   374
      Width           =   1350
   End
   Begin VB.Label Casco 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   150
      Left            =   3195
      TabIndex        =   23
      Top             =   8566
      Width           =   240
   End
   Begin VB.Label Armadura 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   150
      Left            =   2085
      TabIndex        =   22
      Top             =   8551
      Width           =   240
   End
   Begin VB.Label Arma 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   150
      Left            =   960
      TabIndex        =   21
      Top             =   8551
      Width           =   240
   End
   Begin VB.Label Escudo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   150
      Left            =   4350
      TabIndex        =   20
      Top             =   8566
      Width           =   240
   End
   Begin VB.Image Image2 
      Height          =   210
      Left            =   11400
      Top             =   120
      Width           =   270
   End
   Begin VB.Image Image4 
      Height          =   255
      Left            =   11640
      Top             =   120
      Width           =   225
   End
   Begin VB.Image PicMH 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   10320
      Picture         =   "frmMain.frx":6398C
      Stretch         =   -1  'True
      Top             =   2280
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image modoseguro 
      Height          =   254
      Left            =   8820
      Picture         =   "frmMain.frx":6479E
      ToolTipText     =   "Seguro"
      Top             =   583
      Width           =   300
   End
   Begin VB.Label fpps 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   10935
      TabIndex        =   17
      Top             =   120
      Width           =   105
   End
   Begin VB.Label coord3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   10110
      TabIndex        =   16
      Top             =   165
      Width           =   105
   End
   Begin VB.Label Coord2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   9615
      TabIndex        =   15
      Top             =   165
      Width           =   105
   End
   Begin VB.Label Coord 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   165
      Left            =   9150
      TabIndex        =   14
      Top             =   165
      Width           =   105
   End
   Begin VB.Label label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   9000
      TabIndex        =   13
      Top             =   1080
      Width           =   2385
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   10320
      MouseIcon       =   "frmMain.frx":64BDC
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   2760
      Width           =   1245
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8880
      MouseIcon       =   "frmMain.frx":64D2E
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   2760
      Width           =   1365
   End
   Begin VB.Label StaBar 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   9375
      TabIndex        =   10
      Top             =   8460
      Width           =   135
   End
   Begin VB.Label agubar 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   9375
      TabIndex        =   9
      Top             =   7995
      Width           =   135
   End
   Begin VB.Label hambar 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   9375
      TabIndex        =   8
      Top             =   7560
      Width           =   135
   End
   Begin VB.Label ManaBar 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Left            =   9375
      TabIndex        =   7
      Top             =   7110
      Width           =   135
   End
   Begin VB.Label HpBar 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00C0C0FF&
      Height          =   195
      Left            =   9375
      TabIndex        =   6
      Top             =   6675
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   9105
      TabIndex        =   5
      Top             =   1920
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   255
      Index           =   1
      Left            =   10110
      MouseIcon       =   "frmMain.frx":64E80
      MousePointer    =   99  'Custom
      Top             =   6000
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   255
      Index           =   0
      Left            =   10110
      MouseIcon       =   "frmMain.frx":64FD2
      MousePointer    =   99  'Custom
      Top             =   5760
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image cmdInfo 
      Height          =   405
      Left            =   10440
      MouseIcon       =   "frmMain.frx":65124
      MousePointer    =   99  'Custom
      Top             =   5820
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image CmdLanzar 
      Height          =   405
      Left            =   9000
      MouseIcon       =   "frmMain.frx":65276
      MousePointer    =   99  'Custom
      Top             =   5820
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Image InvEqu 
      Height          =   3720
      Left            =   8760
      Picture         =   "frmMain.frx":653C8
      Top             =   2640
      Width           =   2910
   End
   Begin VB.Label GldLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   10560
      TabIndex        =   0
      Top             =   6698
      Width           =   1065
   End
   Begin VB.Image Image3 
      Height          =   315
      Index           =   1
      Left            =   10200
      Top             =   6600
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   0
      Left            =   10320
      MouseIcon       =   "frmMain.frx":6E87B
      MousePointer    =   99  'Custom
      Top             =   6990
      Width           =   1365
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   1
      Left            =   10320
      MouseIcon       =   "frmMain.frx":6E9CD
      MousePointer    =   99  'Custom
      Top             =   7380
      Width           =   1365
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   2
      Left            =   10320
      MouseIcon       =   "frmMain.frx":6EB1F
      MousePointer    =   99  'Custom
      Top             =   7770
      Width           =   1365
   End
   Begin VB.Image Hpshp 
      Height          =   165
      Left            =   8700
      Picture         =   "frmMain.frx":6EC71
      Top             =   6690
      Width           =   1410
   End
   Begin VB.Image MANShp 
      Height          =   165
      Left            =   8700
      Picture         =   "frmMain.frx":6F8E9
      Top             =   7125
      Width           =   1410
   End
   Begin VB.Image COMIDAsp 
      Height          =   165
      Left            =   8700
      Picture         =   "frmMain.frx":70561
      Top             =   7575
      Width           =   1410
   End
   Begin VB.Image AGUAsp 
      Height          =   165
      Left            =   8700
      Picture         =   "frmMain.frx":711D9
      Top             =   8010
      Width           =   1410
   End
   Begin VB.Image STAShp 
      Height          =   165
      Left            =   8700
      Picture         =   "frmMain.frx":71E51
      Top             =   8490
      Width           =   1410
   End
   Begin VB.Shape MainViewShp 
      BorderColor     =   &H00404040&
      BorderStyle     =   0  'Transparent
      Height          =   6225
      Left            =   135
      Top             =   2040
      Width           =   8250
   End
   Begin VB.Menu mnuObj 
      Caption         =   "Objeto"
      Visible         =   0   'False
      Begin VB.Menu mnuTirar 
         Caption         =   "Tirar"
      End
      Begin VB.Menu mnuUsar 
         Caption         =   "Usar"
      End
      Begin VB.Menu mnuEquipar 
         Caption         =   "Equipar"
      End
   End
   Begin VB.Menu mnuNpc 
      Caption         =   "NPC"
      Visible         =   0   'False
      Begin VB.Menu mnuNpcDesc 
         Caption         =   "Descripcion"
      End
      Begin VB.Menu mnuNpcComerciar 
         Caption         =   "Comerciar"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ElDeAhora As Double
Private Diferencia As Double
Private ElDeAntes As Double
Private Empezo As Boolean
Private Minimo As Double
Private Maximo As Double
Private Cont As Byte
Private EstuboDesbalanceado As Long
Private ContEngine As Byte
Private TiempoActual As Long
Private Contador As Integer

Private Lorwik As Boolean
Private Const VK_SNAPSHOT = &H2C

Private Declare Sub keybd_event _
Lib "user32" ( _
ByVal bVk As Byte, _
ByVal bScan As Byte, _
ByVal dwFlags As Long, _
ByVal dwExtraInfo As Long)

Public ActualSecond As Long
Public lastSecond As Long
Public tX As Integer
Public tY As Integer
Public MouseX As Long
Public MouseY As Long
Public MouseBoton As Long
Public MouseShift As Long

Dim gDSB As DirectSoundBuffer
Dim gD As DSBUFFERDESC
Dim gW As WAVEFORMATEX
Dim gFileName As String
Dim dsE As DirectSoundEnum
Dim Pos(0) As DSBPOSITIONNOTIFY
Public IsPlaying As Byte

Dim endEvent As Long

Implements DirectXEvent
Private Sub ActualizadorPosicion_Timer()
    If UserPuedeRefrescar Then
        Call SendData("RPU")
        UserPuedeRefrescar = False
        Beep
    End If
End Sub
Private Sub cmdMoverHechi_Click(Index As Integer)
If hlst.listIndex = -1 Then Exit Sub

Select Case Index
Case 0 'subir
    If hlst.listIndex = 0 Then Exit Sub
Case 1 'bajar
    If hlst.listIndex = hlst.ListCount - 1 Then Exit Sub
End Select

Call SendData("DESPHE" & Index + 1 & "," & hlst.listIndex + 1)

Select Case Index
Case 0 'subir
    hlst.listIndex = hlst.listIndex - 1
Case 1 'bajar
    hlst.listIndex = hlst.listIndex + 1
End Select

End Sub
Private Sub DirectXEvent_DXCallback(ByVal eventid As Long)

End Sub
Private Sub CreateEvent()
     endEvent = DirectX.CreateEvent(Me)
End Sub
Public Sub DibujarMH()
PicMH.Visible = True
End Sub

Public Sub DesDibujarMH()
PicMH.Visible = False
End Sub

Public Sub DibujarSeguro()
PicSeg.Visible = True
End Sub

Public Sub DesDibujarSeguro()
PicSeg.Visible = False
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If endEvent Then
        DirectX.DestroyEvent endEvent
    End If
    If prgRun = True Then
        prgRun = False
        Cancel = 1
    End If
End Sub
Private Sub Image2_Click()
Me.WindowState = vbMinimized
End Sub

Private Sub Image4_Click()
Call SendData("/SALIR")
End Sub

Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HookSurfaceHwnd Me
End Sub
Private Sub Image6_Click()
frmcanjes.Show
End Sub
Private Sub Image7_Click()
Call Shell(App.Path & "\LwK-Radio.EXE", vbNormalFocus)
End Sub

Private Sub img_soporte_Click()
frmRGM.Show , frmMain
img_soporte.Visible = False
End Sub

Private Sub mnuEquipar_Click()
    Call EquiparItem
End Sub
Private Sub mnuNPCComerciar_Click()
    SendData "LC" & tX & "," & tY
    SendData "/COMERCIAR"
End Sub

Private Sub mnuNpcDesc_Click()
    SendData "LC" & tX & "," & tY
End Sub

Private Sub mnuTirar_Click()
    Call TirarItem
End Sub

Private Sub mnuUsar_Click()
    Call UsarItem
End Sub
Private Sub modocombate_Click()
Call SendData("TAB")
                    IScombate = Not IScombate
End Sub
Private Sub modoseguro_Click()
frmMain.PicSeg.Visible = True
Call SendData("/SEG")
End Sub
Private Sub PicSeg_Click()
    frmMain.PicSeg.Visible = False
    Call SendData("/SEG")
End Sub
Private Sub Second_Timer()
    ActualSecond = mid(Time, 7, 2)
    ActualSecond = ActualSecond + 1
    If ActualSecond = lastSecond Then End
    lastSecond = ActualSecond
    If Not DialogosClanes Is Nothing Then DialogosClanes.PassTimer
    
    Lorwik = True
End Sub
'ITEM CONTROL
Private Sub TirarItem()
    If (Inventario.SelectedItem > 0 And Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Or (Inventario.SelectedItem = FLAGORO) Then
        If Inventario.Amount(Inventario.SelectedItem) = 1 Then
            SendData "TI" & Inventario.SelectedItem & "," & 1
        Else
           If Inventario.Amount(Inventario.SelectedItem) > 1 Then
            frmCantidad.Show , frmMain
           End If
        End If
    End If
End Sub

Private Sub AgarrarItem()
    SendData "AG"
End Sub

Private Sub UsarItem()
    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then SendData "USA" & Inventario.SelectedItem
End Sub

Private Sub EquiparItem()
    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then _
        SendData "EQUI" & Inventario.SelectedItem
End Sub
'HECHIZOS CONTROL
Private Sub cmdLanzar_Click()
    If hlst.List(hlst.listIndex) <> "(None)" And UserCanAttack = 1 Then
        Call SendData("HL" & hlst.listIndex + 1)
        Call SendData("KU" & Magia)
        UsaMacro = True
        'UserCanAttack = 0
    End If
End Sub
Private Sub CmdLanzar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UsaMacro = False
    CnTd = 0
End Sub
Private Sub CmdInfo_Click()
    Call SendData("INFS" & hlst.listIndex + 1)
End Sub
'OTROS
Private Sub DespInv_Click(Index As Integer)
    Inventario.ScrollInventory (Index = 0)
End Sub

Private Sub Form_Click()
If Cartel Then Cartel = False
  If Not Comerciando Then
        Call ConvertCPtoTP(MainViewShp.Left, MainViewShp.Top, MouseX, MouseY, tX, tY)

        If MouseShift = 0 Then
            If MouseBoton <> vbRightButton Then
                '[ybarra]
                If UsaMacro Then
                    CnTd = CnTd + 1
                        If CnTd = 3 Then
                            SendData "UMH"
                            CnTd = 0
                        End If
                    UsaMacro = False
                End If
                '[/ybarra]
                If UsingSkill = 0 Then
                    SendData "LC" & tX & "," & tY
                Else
                    frmMain.MousePointer = vbDefault
                    If (UsingSkill = Magia Or UsingSkill = Proyectiles) And UserCanAttack = 0 Then Exit Sub
                    SendData "WLC" & tX & "," & tY & "," & UsingSkill
                    If UsingSkill = Magia Or UsingSkill = Proyectiles Then UserCanAttack = 0
                    UsingSkill = 0
                End If
            Else
                Call AbrirMenuViewPort
            End If
        ElseIf (MouseShift And 1) = 1 Then
            If MouseShift = vbLeftButton Then
                Call SendData("/TELEP YO " & UserMap & " " & tX & " " & tY)
            End If
        End If
    End If
    
End Sub
Private Sub Form_DblClick()
    If Not frmForo.Visible Then
        SendData "RC" & tX & "," & tY
    End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If SendTxt.Visible Or SendCMSTXT.Visible Then Exit Sub
    If (Not SendTxt.Visible) And (Not SendCMSTXT.Visible) Then
 
        If LenB(CustomKeys.ReadableName(KeyCode)) > 0 Then
       
            Select Case KeyCode
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleMusic)
                    If Not Audio.PlayingMusic Then
                        Musica = True
                        
                        Call Extract_File2(Midi, App.Path & "\ARCHIVOS", CStr(currentMidi) & ".mid", Windows_Temp_Dir, False)
                        Audio.PlayMIDI CStr(currentMidi) & ".mid"
                        Delete_File (Windows_Temp_Dir & CStr(currentMidi) & ".mid")
                    Else
                        Musica = False
                        Audio.StopMidi
                    End If
               
                Case CustomKeys.BindedKey(eKeyType.mKeyGetObject)
                    Call AgarrarItem
               
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleCombatMode)
                    Call SendData("TAB")
                    IScombate = Not IScombate
                    
                    If IScombate Then
                    frmMain.combate.Visible = False
                    Else
                    frmMain.combateII.Visible = True
               End If
               
                Case CustomKeys.BindedKey(eKeyType.mKeyEquipObject)
                    Call EquiparItem
               
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleNames)
                    Nombres = Not Nombres
               
                Case CustomKeys.BindedKey(eKeyType.mKeyTamAnimal)
                    Call SendData("KU" & Domar)
               
                Case CustomKeys.BindedKey(eKeyType.mKeySteal)
                    Call SendData("KU" & Robar)
                           
                Case CustomKeys.BindedKey(eKeyType.mKeyHide)
                    Call SendData("KU" & Ocultarse)
               
                Case CustomKeys.BindedKey(eKeyType.mKeyDropObject)
                    Call TirarItem
               
                Case CustomKeys.BindedKey(eKeyType.mKeyUseObject)
                    If Not NoPuedeUsar Then
                        NoPuedeUsar = True
                        Call UsarItem
                    End If
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyQuest)
                Call SendData("QLR")
               
                Case CustomKeys.BindedKey(eKeyType.mKeyRequestRefresh)
                    If UserPuedeRefrescar Then
                        Call SendData("RPU")
                        UserPuedeRefrescar = False
                        Beep
                    End If
                   
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleSafeMode)
                       frmMain.PicSeg.Visible = False
    Call SendData("/SEG")
 
 Case CustomKeys.BindedKey(eKeyType.mKeyMapView)
               FrmMapa.Show , frmMain
 
  Case CustomKeys.BindedKey(eKeyType.mKeyTrabajo)
                If frmMain.WorkMacro.Enabled = True Then
                    frmMain.WorkMacro.Enabled = False
                    Call AddtoRichTextBox(frmMain.RecTxt, "Macro de Trabajo Desactivado.", 255, 255, 255, False, False, False)
                Else
                    frmMain.WorkMacro.Enabled = True
                    Call AddtoRichTextBox(frmMain.RecTxt, "Macro de Trabajo Activado.", 255, 255, 255, False, False, False)
                End If
 
            End Select
        Else
 
        End If
    End If
   
    Select Case KeyCode
        Case CustomKeys.BindedKey(eKeyType.mKeyTalkWithGuild)
                If SendTxt.Visible Then Exit Sub
                If Not frmCantidad.Visible Then
                    SendCMSTXT.Visible = True
                    SendCMSTXT.SetFocus
                End If
       
        Case CustomKeys.BindedKey(eKeyType.mKeyTakeScreenShot)
            Dim i As Integer
                    For i = 1 To 1000
                If Not FileExist(App.Path & "\Fotos\Foto" & i & ".bmp", vbNormal) Then Exit For
                    Next
                    Call Capturar_Guardar(App.Path & "/Fotos/Foto" & i & ".bmp")
                Call AddtoRichTextBox(frmMain.RecTxt, "Foto" & i & ".bmp Guardada en la Carpeta Fotos. Puedes subirla a http://www.lwk-images.com.ar", 255, 255, 255, False, False, False)
           
        Case CustomKeys.BindedKey(eKeyType.mKeyAttack)
            If (UserCanAttack = 1) And _
                   (Not UserDescansar) And _
                   (Not UserMeditar) Then
                        SendData "AT"
                        UserCanAttack = 0
                If IScombate Then
                    charlist(UserCharIndex).Arma.WeaponWalk(charlist(UserCharIndex).Heading).Started = 1
                        charlist(UserCharIndex).Arma.WeaponAttack = GrhData(charlist(UserCharIndex).Arma.WeaponWalk(charlist(UserCharIndex).Heading).GrhIndex).NumFrames + 1
                    Exit Sub
                End If
            End If
       
        Case CustomKeys.BindedKey(eKeyType.mKeyTalk)
                If SendCMSTXT.Visible Then Exit Sub
                If Not frmCantidad.Visible Then
                    SendTxt.Visible = True
                SendTxt.SetFocus
                End If
            Case vbKeyF1:
            Call SecuLwK
                If Lorwik Then
                    Call DoAccionTecla("F1")
                    Lorwik = False
                ElseIf Not Lorwik Then
                    Exit Sub
                End If
            Case vbKeyF2:
            Call SecuLwK
                If Lorwik Then
                    Call DoAccionTecla("F2")
                    Lorwik = False
                ElseIf Not Lorwik Then
                    Exit Sub
                End If
            Case vbKeyF3:
            Call SecuLwK
                If Lorwik Then
                    Call DoAccionTecla("F3")
                    Lorwik = False
                ElseIf Not Lorwik Then
                    Exit Sub
                End If
            Case vbKeyF4:
            Call SecuLwK
                If Lorwik Then
                    Call DoAccionTecla("F4")
                    Lorwik = False
                ElseIf Not Lorwik Then
                    Exit Sub
                End If
            Case vbKeyF5:
            Call SecuLwK
                If Lorwik Then
                    Call DoAccionTecla("F5")
                    Lorwik = False
                ElseIf Not Lorwik Then
                    Exit Sub
                End If
            Case vbKeyF6:
                MsgBox "Acción prohibida"
                Call SecuLwK
           Case vbKeyF7:
                MsgBox "Acción prohibida"
                Call SecuLwK
            Case vbKeyF8:
                MsgBox "Acción prohibida"
                Call SecuLwK
            Case vbKeyF9:
                MsgBox "Acción prohibida"
                Call SecuLwK
            Case vbKeyF10:
                MsgBox "Acción prohibida"
                Call SecuLwK
            Case vbKeyF11:
                MsgBox "Acción prohibida"
                Call SecuLwK
            Case vbKeyF12:
                MsgBox "Acción prohibida"
                Call SecuLwK
                '/Lorwik > Tapamos los posibles agujeros para que no entren las ratas.
                
            Case vbKeyControl:
                If (UserCanAttack = 1) And _
                   (Not UserDescansar) And _
                   (Not UserMeditar) Then
                    SendData "AT"
                    UserCanAttack = 0
                End If
    End Select
End Sub

Private Sub Form_Load()

        With frmMain
        .Width = 12000
        .Height = 9000
    End With

Call Make_Transparent_Richtext(RecTxt.hwnd)

InvEqu.Picture = General_Load_Picture_From_Resource("inventario.gif")
    frmMain.Picture = General_Load_Picture_From_Resource("Todo.gif")
    
   Me.Left = 0
   Me.Top = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseX = X
    MouseY = Y
    exp.Visible = False
lblPorcLvl.Visible = True

GldLbl.Visible = True
GldLbl2.Visible = False
End Sub

Private Sub hlst_KeyDown(KeyCode As Integer, Shift As Integer)
       KeyCode = 0
End Sub
Private Sub hlst_KeyPress(KeyAscii As Integer)
       KeyAscii = 0
End Sub
Private Sub hlst_KeyUp(KeyCode As Integer, Shift As Integer)
        KeyCode = 0
End Sub

Private Sub Image1_Click(Index As Integer)
    Call Audio.PlayWave(SND_CLICK)
    Select Case Index
        Case 0
            '[MatuX] : 01 de Abril del 2002
                Call frmOpciones.Show(vbModeless, frmMain)
            '[END]
        Case 1
            LlegaronAtrib = False
            LlegaronSkills = False
            LlegoFama = False
            SendData "ATRI"
            SendData "ESKI"
            SendData "FEST"
            SendData "FAMA"
            Do While Not LlegaronSkills Or Not LlegaronAtrib Or Not LlegoFama
                DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
            Loop
            frmEstadisticas.Iniciar_Labels
            frmEstadisticas.Show , frmMain
            LlegaronAtrib = False
            LlegaronSkills = False
            LlegoFama = False
        Case 2
            If Not frmGuildLeader.Visible Then _
                Call SendData("GLINFO")
    End Select
End Sub
Private Sub Image3_Click(Index As Integer)
frmDarOro.Show , frmMain
End Sub
Private Sub Label1_Click()
    Dim i As Integer
    For i = 1 To NUMSKILLS
        frmSkills3.Text1(i).Caption = UserSkills(i)
    Next i
    Alocados = SkillPoints
    frmSkills3.puntos.Caption = SkillPoints
    frmSkills3.Show , frmMain
End Sub
Private Sub Label4_Click()
    Call Audio.PlayWave(SND_CLICK)

   InvEqu.Picture = General_Load_Picture_From_Resource("inventario.gif")

    'DespInv(0).Visible = True
    'DespInv(1).Visible = True
    picInv.Visible = True

    hlst.Visible = False
    cmdInfo.Visible = False
    CmdLanzar.Visible = False
    
    cmdMoverHechi(0).Visible = True
    cmdMoverHechi(1).Visible = True
    ItemName.Visible = True
End Sub

Private Sub Label7_Click()
    Call Audio.PlayWave(SND_CLICK)

      InvEqu.Picture = General_Load_Picture_From_Resource("hechizos.gif")
    '%%%%%%OCULTAMOS EL INV&&&&&&&&&&&&
    'DespInv(0).Visible = False
    'DespInv(1).Visible = False
    picInv.Visible = False
    hlst.Visible = True
    cmdInfo.Visible = True
    CmdLanzar.Visible = True
    
    cmdMoverHechi(0).Visible = True
    cmdMoverHechi(1).Visible = True
    ItemName.Visible = False
End Sub
Private Sub picInv_DblClick()
    If frmCarp.Visible Or frmHerrero.Visible Then Exit Sub
    Call UsarItem
    Call EquiparItem
End Sub
Private Sub picInv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Audio.PlayWave(SND_CLICK)
End Sub
Private Sub RecTxt_Change()
    On Error Resume Next  'el .SetFocus causaba errores al salir y volver a entrar
    If SendTxt.Visible Then
        SendTxt.SetFocus
    ElseIf Me.SendCMSTXT.Visible Then
        SendCMSTXT.SetFocus
    Else
      If (Not frmComerciar.Visible) And _
         (Not frmSkills3.Visible) And _
         (Not frmForo.Visible) And _
         (Not frmEstadisticas.Visible) And _
         (Not frmCantidad.Visible) And _
         (picInv.Visible) Then
            picInv.SetFocus
      End If
    End If
    On Error GoTo 0
End Sub
Private Sub RecTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If picInv.Visible Then
        picInv.SetFocus
    Else
        hlst.SetFocus
    End If
End Sub
Private Sub SecureLwK_Timer()
'Lorwik - Esto lo puse aqui por que me daba pena crear otro timer xD
    Call BuscarEngine
    
    If logged Then
If Not logged Then Exit Sub
    If GetTickCount - TiempoActual > 110 Or GetTickCount - TiempoActual < 109 Then
        Contador = Contador + 1
    Else
        Contador = 0
    End If

End If

    Call SecuLwK
    
    Erase Valores
nProcesos = 0
EnumWindows AddressOf Listar_Ventanas, 0

If sGM.Visible = True Then
    AddtoRichTextBox frmMain.RecTxt, "¡¡ATENCIÓN, LOS ADMINISTRADORES HAN RESPONDIDO TU CONSULTA!!", 252, 151, 53, 1, 0
    AddtoRichTextBox frmMain.RecTxt, "¡¡ATENCIÓN, LOS ADMINISTRADORES HAN RESPONDIDO TU CONSULTA!!", 252, 151, 53, 1, 0
End If
End Sub

Private Sub SendTxt_Change()
    If Len(SendTxt.Text) > 160 Then
        stxtbuffer = "Soy un cheater, avisenle a un gm"
    Else
        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim i As Long
        Dim tempstr As String
        Dim CharAscii As Integer
        
        For i = 1 To Len(SendTxt.Text)
            CharAscii = Asc(mid$(SendTxt.Text, i, 1))
            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                tempstr = tempstr & Chr$(CharAscii)
            End If
        Next i
        
        If tempstr <> SendTxt.Text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            SendTxt.Text = tempstr
        End If
        
        stxtbuffer = SendTxt.Text
    End If
End Sub
Private Sub SendTxt_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub
Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    If KeyCode = vbKeyReturn Then
        If Left$(stxtbuffer, 1) = "/" Then
            If UCase(Left$(stxtbuffer, 8)) = "/PASSWD " Then
                    Dim j As String

                    j = Right$(stxtbuffer, Len(stxtbuffer) - 8)

                    stxtbuffer = "/PASSWD " & j
                    
                                ElseIf UCase$(stxtbuffer) = "/HACERLWK" Then
                frmConsolaTorneo.Show vbModeless, Me
                stxtbuffer = ""
                SendTxt.Text = ""
                KeyCode = 0
                SendTxt.Visible = False
                Exit Sub
                
            ElseIf UCase$(stxtbuffer) = "/FUNDARCLAN" Then
                frmEligeAlineacion.Show vbModeless, Me
                stxtbuffer = ""
                SendTxt.Text = ""
                KeyCode = 0
                SendTxt.Visible = False
                
                Exit Sub
            End If
            Call SendData(stxtbuffer)
    
       'Shout
        ElseIf Left$(stxtbuffer, 1) = "-" Then
            Call SendData("-" & Right$(stxtbuffer, Len(stxtbuffer) - 1))

       'Global
        ElseIf Left$(stxtbuffer, 1) = ";" Then
            Call SendData(":" & Right$(stxtbuffer, Len(stxtbuffer) - 1))

        'Whisper
        ElseIf Left$(stxtbuffer, 1) = "\" Then
            Call SendData("\" & Right$(stxtbuffer, Len(stxtbuffer) - 1))
            Call SendData("@" & Right$(stxtbuffer, Len(stxtbuffer) - 1))

        'Say
        ElseIf stxtbuffer <> "" Then
            Call SendData(";" & stxtbuffer)

        End If

        stxtbuffer = ""
        SendTxt.Text = ""
        KeyCode = 0
        SendTxt.Visible = False
    End If
End Sub
Private Sub SendCMSTXT_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    If KeyCode = vbKeyReturn Then
        'Say
        If stxtbuffercmsg <> "" Then
            Call SendData("/CMSG " & stxtbuffercmsg)
        End If

        stxtbuffercmsg = ""
        SendCMSTXT.Text = ""
        KeyCode = 0
        Me.SendCMSTXT.Visible = False
    End If
End Sub
Private Sub SendCMSTXT_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub
Private Sub SendCMSTXT_Change()
    If Len(SendCMSTXT.Text) > 160 Then
        stxtbuffercmsg = "Soy un cheater, avisenle a un GM"
    Else
        stxtbuffercmsg = SendCMSTXT.Text
    End If
End Sub
'SOCKET1
#If UsarWrench = 1 Then

Private Sub sGM_Click()
frmRGM.Show , frmMain
sGM.Visible = False
End Sub

Private Sub Socket1_Connect()
    Second.Enabled = True
    UsersID = 33
    Call SendData("gIvEmEvAlcOde")
    
End Sub
Private Sub Socket1_Disconnect()
    lastSecond = 0
    Second.Enabled = False
    logged = False
    Connected = False
    
    Socket1.Cleanup
    
    frmConnect.MousePointer = vbDefault
    
    If frmPasswdSinPadrinos.Visible = True Then frmPasswdSinPadrinos.Visible = False
    frmCrearPersonaje.Visible = False
    frmConnect.Visible = True
    
    On Local Error Resume Next
    For i = 0 To Forms.Count - 1
        If Forms(i).Name <> Me.Name And Forms(i).Name <> frmConnect.Name Then
            Unload Forms(i)
        End If
    Next i
    On Local Error GoTo 0
    
    frmMain.Visible = False

    pausa = False
    UserMeditar = False

    UserClase = ""
    UserSexo = ""
    UserRaza = ""
    UserEmail = ""
    
    For i = 1 To NUMSKILLS
        UserSkills(i) = 0
    Next i

    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i

    SkillPoints = 0
    Alocados = 0

    Dialogos.UltimoDialogo = 0
    Dialogos.CantidadDialogos = 0

End Sub
Private Sub Socket1_LastError(ErrorCode As Integer, ErrorString As String, Response As Integer)
    'Handle socket errors
    If ErrorCode = 24036 Then
        Call MsgBox("Por favor espere, intentando completar conexion.", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
        Exit Sub
    End If
    
    Call MsgBox(ErrorString, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
    frmConnect.MousePointer = vbDefault
    Response = 0
    lastSecond = 0
    Second.Enabled = False

    frmMain.Socket1.Disconnect
    
    If Not frmCrearPersonaje.Visible Then
      '  If Not frmBorrar.Visible And Not frmRecuperar.Visible Then
            frmConnect.Show
       ' End If
    Else
        frmCrearPersonaje.MousePointer = vbDefault
    End If
End Sub
Private Sub Socket1_Read(DataLength As Integer, IsUrgent As Integer)
    Dim loopc As Integer

    Dim RD As String
    Dim rBuffer(1 To 500) As String
    Static TempString As String

    Dim CR As Integer
    Dim tChar As String
    Dim sChar As Integer
    Dim Echar As Integer
    Dim aux$
    Dim nfile As Integer

    
Socket1.Read RD, DataLength

    'Check for previous broken data and add to current data
    If TempString <> "" Then
        RD = TempString & RD
        TempString = ""
    End If

    'Check for more than one line
    sChar = 1
    For loopc = 1 To Len(RD)

        tChar = mid$(RD, loopc, 1)

        If tChar = ENDC Then
            CR = CR + 1
            Echar = loopc - sChar
            rBuffer(CR) = mid$(RD, sChar, Echar)
            sChar = loopc + 1
        End If

    Next loopc

    'Check for broken line and save for next time
    If Len(RD) - (sChar - 1) <> 0 Then
        TempString = mid$(RD, sChar, Len(RD))
    End If

    'Send buffer to Handle data
    For loopc = 1 To CR
        Call HandleData(rBuffer(loopc))
    Next loopc
End Sub
#End If
Private Sub AbrirMenuViewPort()
#If (ConMenuseConextuales = 1) Then

If tX >= MinXBorder And tY >= MinYBorder And _
    tY <= MaxYBorder And tX <= MaxXBorder Then
    If MapData(tX, tY).CharIndex > 0 Then
        If charlist(MapData(tX, tY).CharIndex).invisible = False Then
        
            Dim i As Long
            Dim m As New frmMenuseFashion
            
            Load m
            m.SetCallback Me
            m.SetMenuId 1
            m.ListaInit 2, False
            
            If charlist(MapData(tX, tY).CharIndex).Nombre <> "" Then
                m.ListaSetItem 0, charlist(MapData(tX, tY).CharIndex).Nombre, True
            Else
                m.ListaSetItem 0, "<NPC>", True
            End If
            m.ListaSetItem 1, "Comerciar"
            
            m.ListaFin
            m.Show , Me

        End If
    End If
End If

#End If
End Sub
Public Sub CallbackMenuFashion(ByVal MenuId As Long, ByVal Sel As Long)
Select Case MenuId

Case 0 'Inventario
    Select Case Sel
    Case 0
    Case 1
    Case 2 'Tirar
        Call TirarItem
    Case 3 'Usar
        If Not NoPuedeUsar Then
            NoPuedeUsar = True
            Call UsarItem
        End If
    Case 3 'equipar
        Call EquiparItem
    End Select
    
Case 1 'Menu del ViewPort del engine
    Select Case Sel
    Case 0 'Nombre
        SendData "LC" & tX & "," & tY
    Case 1 'Comerciar
        Call SendData("LC" & tX & "," & tY)
        Call SendData("/COMERCIAR")
    End Select
End Select
End Sub
Private Sub Minimap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then Call SendData("/TELEP YO " & UserMap & " " & CByte(X) & " " & CByte(Y))
End Sub
Private Sub lblPorcLvl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     MouseX = X
     MouseY = Y
        lblPorcLvl.Visible = False
        exp.Visible = True
        
End Sub
 Private Sub GldLbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     MouseX = X
     MouseY = Y
        GldLbl.Visible = False
        GldLbl2.Visible = True
End Sub
Private Sub Macros_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbRightButton Then
    frmMacros.Show vbModeless, frmMain
Else
    If Lorwik Then
        Call DoAccionTecla("F" & Index)
        Lorwik = False
    ElseIf Not Lorwik Then
        Exit Sub
    End If
End If
End Sub
Private Sub WorkMacro_Timer()

If Me.ItemName.Caption = "Hacha de Leñador" Or Me.ItemName.Caption = "Piquete de Minero" Or Me.ItemName.Caption = "Caña de Pescar" Or Me.ItemName.Caption = "Minerales de Hierro" Or Me.ItemName.Caption = "Minerales de Plata" Or Me.ItemName.Caption = "Minerales de Oro" Or Me.ItemName.Caption = "Red de Pesca " Then
    SendData "USA" & Inventario.SelectedItem
    SendData "WLC" & tX & "," & tY & "," & UsingSkill
Else
    AddtoRichTextBox frmMain.RecTxt, "No Puedes Usar el Macro Con Este item!", 255, 255, 255, False, False, False
    frmMain.WorkMacro.Enabled = False
    Call AddtoRichTextBox(frmMain.RecTxt, "Macro de Trabajo Desactivado.", 255, 255, 255, False, False, False)
    Exit Sub
End If

End Sub
