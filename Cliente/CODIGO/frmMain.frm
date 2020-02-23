VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   345
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
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox MainViewPic 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   6015
      Left            =   180
      MousePointer    =   99  'Custom
      ScaleHeight     =   401
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   544
      TabIndex        =   25
      Top             =   2100
      Width           =   8160
      Begin VB.Timer macrotrabajo 
         Enabled         =   0   'False
         Interval        =   700
         Left            =   7680
         Top             =   480
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   7680
         Top             =   60
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
   End
   Begin VB.ListBox LComm 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1005
      ItemData        =   "frmMain.frx":3AFA
      Left            =   165
      List            =   "frmMain.frx":3AFC
      Sorted          =   -1  'True
      TabIndex        =   23
      Top             =   540
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox SendTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   180
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1695
      Visible         =   0   'False
      Width           =   8160
   End
   Begin VB.PictureBox Minimap 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1395
      Left            =   6840
      ScaleHeight     =   93
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   16
      Top             =   180
      Width           =   1500
      Begin VB.Shape UserArea 
         BorderColor     =   &H80000004&
         Height          =   225
         Left            =   615
         Top             =   600
         Width           =   300
      End
      Begin VB.Shape UserM 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   75
         Left            =   750
         Shape           =   3  'Circle
         Top             =   690
         Width           =   60
      End
   End
   Begin VB.ListBox hlst 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2565
      Left            =   8880
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2370
      Visible         =   0   'False
      Width           =   2550
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
      Height          =   2520
      Left            =   8910
      ScaleHeight     =   168
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   167
      TabIndex        =   0
      Top             =   2400
      Width           =   2505
   End
   Begin RichTextLib.RichTextBox RecTxt 
      Height          =   1380
      Left            =   180
      TabIndex        =   24
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   180
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   2434
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      OLEDropMode     =   0
      TextRTF         =   $"frmMain.frx":3AFE
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
   Begin VB.Image Image7 
      Height          =   480
      Left            =   11340
      Picture         =   "frmMain.frx":3B7B
      Top             =   720
      Width           =   480
   End
   Begin VB.Image Command1 
      Height          =   255
      Left            =   10320
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Label lblMapName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mapa Desconocido"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8640
      TabIndex        =   28
      Top             =   8295
      Width           =   3015
   End
   Begin VB.Image Miniminizar 
      Height          =   255
      Left            =   11400
      Top             =   285
      Width           =   255
   End
   Begin VB.Image Cerrar 
      Height          =   255
      Left            =   11640
      Top             =   120
      Width           =   255
   End
   Begin VB.Image Clima 
      Height          =   480
      Left            =   6600
      Top             =   8310
      Width           =   1695
   End
   Begin VB.Label GldLbl 
      Alignment       =   2  'Center
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
      Height          =   195
      Left            =   10680
      TabIndex        =   27
      Top             =   5760
      Width           =   1110
   End
   Begin VB.Label lblFPS 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   5100
      TabIndex        =   26
      Top             =   8460
      Width           =   555
   End
   Begin VB.Image cmdQuests 
      Height          =   255
      Left            =   10320
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Label ItemInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "(nada)"
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
      Height          =   240
      Left            =   8595
      TabIndex        =   22
      Top             =   5130
      Width           =   3150
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   255
      Index           =   1
      Left            =   11520
      MouseIcon       =   "frmMain.frx":4445
      MousePointer    =   99  'Custom
      Top             =   2700
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   255
      Index           =   0
      Left            =   11520
      MouseIcon       =   "frmMain.frx":4597
      MousePointer    =   99  'Custom
      Top             =   3000
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image PicResu 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   5925
      Stretch         =   -1  'True
      Top             =   8385
      Width           =   375
   End
   Begin VB.Label lblWeapon 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3480
      TabIndex        =   21
      Top             =   8475
      Width           =   855
   End
   Begin VB.Label LblShield 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   20
      Top             =   8475
      Width           =   615
   End
   Begin VB.Label lblHelm 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   19
      Top             =   8475
      Width           =   375
   End
   Begin VB.Label lblArmor 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   18
      Top             =   8475
      Width           =   540
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
      Height          =   555
      Left            =   8880
      MouseIcon       =   "frmMain.frx":46E9
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   1800
      Width           =   1245
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
      Left            =   10200
      MouseIcon       =   "frmMain.frx":483B
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   1875
      Width           =   1245
   End
   Begin VB.Image cmdInfo 
      Height          =   405
      Left            =   10560
      MouseIcon       =   "frmMain.frx":498D
      MousePointer    =   99  'Custom
      Top             =   5040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image CmdLanzar 
      Height          =   405
      Left            =   8760
      MouseIcon       =   "frmMain.frx":4ADF
      MousePointer    =   99  'Custom
      Top             =   5040
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.Label lblVida 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   8880
      TabIndex        =   13
      Top             =   5970
      UseMnemonic     =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblMana 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   8880
      TabIndex        =   12
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Label lblHambre 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   8880
      TabIndex        =   11
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Label lblSed 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   8880
      TabIndex        =   10
      Top             =   7110
      Width           =   1095
   End
   Begin VB.Label lblEnergia 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   8880
      TabIndex        =   9
      Top             =   7500
      Width           =   1095
   End
   Begin VB.Label lblStrg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   10230
      TabIndex        =   8
      Top             =   7905
      Width           =   210
   End
   Begin VB.Label lblDext 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   9675
      TabIndex        =   7
      Top             =   7920
      Width           =   210
   End
   Begin VB.Label lblPorcLvl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0.0%"
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
      Left            =   9960
      TabIndex        =   6
      Top             =   1275
      Width           =   465
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
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
      Height          =   195
      Left            =   9480
      TabIndex        =   5
      Top             =   660
      Width           =   2265
   End
   Begin VB.Label LvlLbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8880
      TabIndex        =   4
      Top             =   405
      Width           =   210
   End
   Begin VB.Label exp 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   9795
      TabIndex        =   3
      Top             =   1005
      Width           =   285
   End
   Begin VB.Image Image3 
      Height          =   315
      Index           =   0
      Left            =   10350
      Top             =   5640
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   0
      Left            =   10320
      MouseIcon       =   "frmMain.frx":4C31
      MousePointer    =   99  'Custom
      Top             =   6480
      Width           =   1410
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   1
      Left            =   10320
      MouseIcon       =   "frmMain.frx":4D83
      MousePointer    =   99  'Custom
      Top             =   6720
      Width           =   1365
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   2
      Left            =   10320
      MouseIcon       =   "frmMain.frx":4ED5
      MousePointer    =   99  'Custom
      Top             =   7440
      Width           =   1410
   End
   Begin VB.Image STAShp 
      Height          =   165
      Left            =   8700
      Top             =   7485
      Width           =   1410
   End
   Begin VB.Image AGUAsp 
      Height          =   165
      Left            =   8700
      Top             =   7110
      Width           =   1410
   End
   Begin VB.Image COMIDAsp 
      Height          =   165
      Left            =   8700
      Top             =   6720
      Width           =   1410
   End
   Begin VB.Image MANShp 
      Height          =   165
      Left            =   8700
      Top             =   6375
      Width           =   1410
   End
   Begin VB.Image Hpshp 
      Height          =   165
      Left            =   8700
      Top             =   5970
      Width           =   1410
   End
   Begin VB.Image InvEqu 
      Height          =   3735
      Left            =   8520
      Top             =   1680
      Width           =   3300
   End
   Begin VB.Image ExpShp 
      Height          =   195
      Left            =   8700
      Top             =   1275
      Width           =   2925
   End
   Begin VB.Label Coord 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "[000, 00, 00]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   9480
      TabIndex        =   2
      Top             =   8490
      Width           =   1335
   End
   Begin VB.Image BarraFuerza 
      Height          =   150
      Left            =   10125
      Top             =   7920
      Width           =   405
   End
   Begin VB.Image BarraAgilidad 
      Height          =   150
      Left            =   9555
      Top             =   7920
      Width           =   405
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private KeyFisico As Boolean

Public tX As Byte
Public tY As Byte
Public MouseX As Long
Public MouseY As Long
Public MouseBoton As Long
Public MouseShift As Long
Private clicX As Long
Private clicY As Long

Public IsPlaying As Byte

Private m_Jpeg As clsJpeg
Private m_FileName As String
Public Attack As Boolean

Private Sub Cerrar_Click()
'Lorwik> Cerramos correctamente el cliente...
Call General_Set_Wav(SND_CLICK)
Call CloseClient
End Sub

Private Sub Command1_Click()
    If Not frmParty.Visible = True Then
        Call WriteRequestPartyForm
    End If
End Sub

Private Sub Image7_Click()
Call Shell(App.Path & "\RadioXtreme.EXE", vbNormalFocus)
End Sub

Private Sub MainViewPic_Click()
    form_click
End Sub

Private Sub MainViewPic_DblClick()
    Form_DblClick
End Sub

Private Sub MainViewPic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub MainViewPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseX = X
    MouseY = Y
End Sub

Private Sub MainViewPic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    clicX = X
    clicY = Y
End Sub

Private Sub Minimap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then Call ParseUserCommand("/TELEP YO " & UserMap & " " & CByte(X) & " " & CByte(Y))
End Sub
Private Sub cmdMoverHechi_Click(Index As Integer)
    If hlst.ListIndex = -1 Then Exit Sub
    Dim sTemp As String

    Select Case Index
        Case 1 'subir
            If hlst.ListIndex = 0 Then Exit Sub
        Case 0 'bajar
            If hlst.ListIndex = hlst.ListCount - 1 Then Exit Sub
    End Select

    Call WriteMoveSpell(Index, hlst.ListIndex + 1)
    
    Select Case Index
        Case 1 'subir
            sTemp = hlst.List(hlst.ListIndex - 1)
            hlst.List(hlst.ListIndex - 1) = hlst.List(hlst.ListIndex)
            hlst.List(hlst.ListIndex) = sTemp
            hlst.ListIndex = hlst.ListIndex - 1
        Case 0 'bajar
            sTemp = hlst.List(hlst.ListIndex + 1)
            hlst.List(hlst.ListIndex + 1) = hlst.List(hlst.ListIndex)
            hlst.List(hlst.ListIndex) = sTemp
            hlst.ListIndex = hlst.ListIndex + 1
    End Select
End Sub

Public Sub ControlSeguroResu(ByVal Mostrar As Boolean)
If Mostrar Then
    If Not PicResu.Visible Then
        PicResu.Picture = General_Load_Picture_From_Resource("89.gif")
    End If
Else
    If PicResu.Visible Then
        PicResu.Picture = General_Load_Picture_From_Resource("79.gif")
    End If
End If
End Sub

Private Sub cmdQuests_Click()
Call General_Set_Wav(SND_CLICK)
'Maneja el evento click del CommandButton cmdQuests.
    Call WriteQuestListRequest
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If Not KeyFisico Then Exit Sub
    If (Not SendTxt.Visible) Then
        
        'Checks if the key is valid
        If LenB(CustomKeys.ReadableName(KeyCode)) > 0 Then
            Select Case KeyCode
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleMusic)
                    If Audio.MusicActivated = True Then
                        Audio.MusicActivated = False
                        Call WriteVar(App.Path & "\Init\Config.cfg", "Sound", "MP3", 0)
                    Else
                        Audio.MusicActivated = True
                        Call WriteVar(App.Path & "\Init\Config.cfg", "Sound", "MP3", 1)
                    End If
                
                Case CustomKeys.BindedKey(eKeyType.mKeyGetObject)
                    Call AgarrarItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleCombatMode)
                    Call WriteCombatModeToggle
                    IScombate = Not IScombate
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyEquipObject)
                    Call EquiparItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleNames)
                    Nombres = Not Nombres
                    
                 Case CustomKeys.BindedKey(eKeyType.mKeyMap)
                    frmMapa.Show , frmMain
                
                Case CustomKeys.BindedKey(eKeyType.mKeyTamAnimal)
                    If UserEstado = 1 Then
                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                        End With
                    Else
                        Call WriteWork(eSkill.Domar)
                    End If
                    
                Case CustomKeys.BindedKey(eKeyType.mKeySteal)
                    If UserEstado = 1 Then
                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                        End With
                    Else
                        Call WriteWork(eSkill.Robar)
                    End If
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyHide)
                    If UserEstado = 1 Then
                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                        End With
                    Else
                        Call WriteWork(eSkill.Ocultarse)
                    End If
                                    
                Case CustomKeys.BindedKey(eKeyType.mKeyDropObject)
                    Call TirarItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyUseObject)
                        If MainTimer.Check(TimersIndex.UseItemWithU) Then Call UsarItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyRequestRefresh)
                    If MainTimer.Check(TimersIndex.SendRPU) Then
                        Call WriteRequestPositionUpdate
                        Beep
                    End If
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleResuscitationSafe)
                    Call WriteResuscitationToggle
            End Select
        End If
    End If
    
    Select Case KeyCode
        Case CustomKeys.BindedKey(eKeyType.mKeyTalkWithGuild)
            If (Not frmComerciar.Visible) And (Not frmComerciarUsu.Visible) And _
              (Not frmBancoObj.Visible) And _
              (Not frmMSG.Visible) And (Not frmForo.Visible) And _
              (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) Then
                SendTxt.Text = "/CMSG "
                SendTxt.Visible = True
                SendTxt.SetFocus
            End If
        
        Case CustomKeys.BindedKey(eKeyType.mKeyTakeScreenShot)
            Call frmMain.Client_Screenshot(frmMain.hDC, 800, 600)
            
        Case CustomKeys.BindedKey(eKeyType.mKeyAttack)
                If Shift <> 0 Then Exit Sub
   
                If Not MainTimer.Check(TimersIndex.Arrows, False) Then Exit Sub 'Check if arrows interval has finished.
                If Not MainTimer.Check(TimersIndex.CastSpell, False) Then 'Check if spells interval has finished.
                    If Not MainTimer.Check(TimersIndex.CastAttack) Then Exit Sub 'Corto intervalo Golpe-Hechizo
                Else
                    If Not MainTimer.Check(TimersIndex.Attack) Or UserDescansar Or UserMeditar Then Exit Sub
                End If
                
               Call WriteAttack
               Attack = True

        
        Case CustomKeys.BindedKey(eKeyType.mKeyTalk)
            If (Not frmComerciar.Visible) And (Not frmComerciarUsu.Visible) And _
              (Not frmBancoObj.Visible) And _
              (Not frmMSG.Visible) And (Not frmForo.Visible) And _
              (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) Then
                SendTxt.Visible = True
                SendTxt.SetFocus
            End If
            
            Case vbKeyF1:
                Call DoAccionTecla("F1")
                Call SecuLwK
                
            Case vbKeyF2:
                Call DoAccionTecla("F2")
                Call SecuLwK

            Case vbKeyF3:
                Call DoAccionTecla("F3")
                Call SecuLwK

            Case vbKeyF4:
                Call DoAccionTecla("F4")
                Call SecuLwK
                
            Case vbKeyF5:
                Call DoAccionTecla("F5")
                Call SecuLwK

            Case vbKeyF6:
                Call DoAccionTecla("F6")
                Call SecuLwK
                
           Case vbKeyF7:
                Call DoAccionTecla("F7")
                Call SecuLwK
                
            Case vbKeyF8:
                Call DoAccionTecla("F8")
                Call SecuLwK
                
            Case vbKeyF9:
                Call DoAccionTecla("F9")
                Call SecuLwK
                
            Case vbKeyF10:
                Call DoAccionTecla("F10")
                Call SecuLwK
                
            Case vbKeyF11:
                Call DoAccionTecla("F11")
                Call SecuLwK
                
            Case vbKeyF12:
                Call DoAccionTecla("F12")
                Call SecuLwK
    End Select
    KeyFisico = False
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = vbLeftButton) Then Call Auto_Drag(Me.hwnd)
    
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    clicX = X
    clicY = Y
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If prgRun = True Then
        prgRun = False
        Cancel = 1
    End If
End Sub
'[Seguridad LwK - AntiMacros de palo]
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'********************************************u
'Lorwik> El codigo no es completamente mio :(
'********************************************

If Not GetAsyncKeyState(KeyCode) < 0 Then Exit Sub
KeyFisico = True
End Sub
'[/Seguridad LwK - AntiMacros de palo]

Private Sub mnuEquipar_Click()
    Call EquiparItem
End Sub

Private Sub mnuNPCComerciar_Click()
    Call WriteLeftClick(tX, tY)
    Call WriteCommerceStart
End Sub

Private Sub mnuNpcDesc_Click()
    Call WriteLeftClick(tX, tY)
End Sub

Private Sub mnuTirar_Click()
    Call TirarItem
End Sub

Private Sub mnuUsar_Click()
    Call UsarItem
End Sub
Private Sub Coord_Click()
    AddtoRichTextBox frmMain.RecTxt, "Estas coordenadas son tu ubicación en el mapa. Utiliza la letra L para corregirla si esta no se corresponde con la del servidor por efecto del Lag.", 255, 255, 255, False, False, False
End Sub


Private Sub Miniminizar_Click()
Call General_Set_Wav(SND_CLICK)
'Lorwik> Miniminizamos bien el cliente, dejandolo en la barra de herramientas :P
Me.WindowState = vbMinimized
End Sub

Private Sub Seguridad_Timer()
'*********************************
'Hace comprobaciones de seguridad
'Author: Manuel (Lorwik)
'*********************************

    '******Anti Cheats & Macros*****
    'Call Externos("Cheat")
    'Call Externos("Macro")
    'Call Externos("Makro")
    
    '*******Anti Speed Hack*********
    If AntiSh(FramesPerSecCounter) Then
        Call AntiShOn
        End
    End If
End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)

    If Opciones.AutoComandos = False Then
       'Si pulso "/" para ingresar un comando
        If KeyCode = "111" Then
            'Muestro el listado de los comandos
            LComm.Visible = True
            Mostrando = True
        End If
    End If

    If Mostrando Then
        SendTxt.SelStart = Len(SendTxt)
       
        'Buscamos el comando
        SearchCommand SendTxt.Text
   End If
   
   If Opciones.AutoComandos = False Then
    If KeyCode = vbKeySpace Then
         Mostrando = False
         LComm.Visible = False
    End If
   End If
    'Send text
    If KeyCode = vbKeyReturn Then
        If LenB(stxtbuffer) <> 0 Then Call ParseUserCommand(stxtbuffer)
        
        stxtbuffer = ""
        SendTxt.Text = ""
        KeyCode = 0
        SendTxt.Visible = False
        If Opciones.AutoComandos = False Then
            LComm.Visible = False
        End If
    End If
End Sub

'[END]'

''''''''''''''''''''''''''''''''''''''
'     ITEM CONTROL                   '
''''''''''''''''''''''''''''''''''''''

Private Sub TirarItem()
    If UserEstado = 1 Then
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
        End With
    Else
        If (Inventario.SelectedItem > 0 And Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Or (Inventario.SelectedItem = FLAGORO) Then
            If Inventario.Amount(Inventario.SelectedItem) = 1 Then
                Call WriteDrop(Inventario.SelectedItem, 1)
            Else
                If Inventario.Amount(Inventario.SelectedItem) > 1 Then
                frmCantidad.Show , frmMain
                End If
            End If
        End If
    End If
End Sub

Private Sub AgarrarItem()
    If UserEstado = 1 Then
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
        End With
    Else
        Call WritePickUp
    End If
End Sub

Private Sub UsarItem()
    If pausa Then Exit Sub
    
    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then _
        Call WriteUseItem(Inventario.SelectedItem)
End Sub

Private Sub EquiparItem()
    If UserEstado = 1 Then
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
        End With
    Else
        If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then _
        Call WriteEquipItem(Inventario.SelectedItem)
    End If
End Sub
Private Sub cmdLanzar_Click()
    If hlst.List(hlst.ListIndex) <> "(None)" And MainTimer.Check(TimersIndex.Work, False) Then
        If UserEstado = 1 Then
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
            End With
        Else
            Call WriteCastSpell(hlst.ListIndex + 1)
            Call WriteWork(eSkill.Magia)
            UsaMacro = True
        End If
    End If
End Sub

Private Sub CmdLanzar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UsaMacro = False
    CnTd = 0
End Sub

Private Sub cmdINFO_Click()
    If hlst.ListIndex <> -1 Then
        Call WriteSpellInfo(hlst.ListIndex + 1)
    End If
End Sub

Private Sub DespInv_Click(Index As Integer)
    Inventario.ScrollInventory (Index = 0)
End Sub

Private Sub form_click()
    SendTxt.Visible = False
    If Cartel Then Cartel = False

    If Not Comerciando Then
        Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
         
        If MouseShift = 0 Then
            If MouseBoton <> vbRightButton Then
                '[ybarra]
                If UsaMacro Then
                    CnTd = CnTd + 1
                    If CnTd = 3 Then
                        Call WriteUseSpellMacro
                        CnTd = 0
                    End If
                    UsaMacro = False
                End If
                '[/ybarra]
                If UsingSkill = 0 Then
                    Call WriteLeftClick(tX, tY)
                Else
                    
                    If Not MainTimer.Check(TimersIndex.Arrows, False) Then 'Check if arrows interval has finished.
                        frmMain.MousePointer = vbDefault
                        UsingSkill = 0
                        With FontTypes(FontTypeNames.FONTTYPE_TALK)
                            Call AddtoRichTextBox(frmMain.RecTxt, "No puedes lanzar proyectiles tan rápido.", .red, .green, .blue, .bold, .italic)
                        End With
                        Exit Sub
                    End If
                    
                    'Splitted because VB isn't lazy!
                    If UsingSkill = Proyectiles Then
                        If Not MainTimer.Check(TimersIndex.Arrows) Then
                            frmMain.MousePointer = vbDefault
                            UsingSkill = 0
                            With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                Call AddtoRichTextBox(frmMain.RecTxt, "No puedes lanzar proyectiles tan rápido.", .red, .green, .blue, .bold, .italic)
                            End With
                            Exit Sub
                        End If
                    End If
                    
                    'Splitted because VB isn't lazy!
                    If (UsingSkill = Pesca Or UsingSkill = Robar Or UsingSkill = Talar Or UsingSkill = Mineria Or UsingSkill = FundirMetal) Then
                        If Not MainTimer.Check(TimersIndex.Work) Then
                            frmMain.MousePointer = vbDefault
                            UsingSkill = 0
                            Exit Sub
                        End If
                    End If
                    
                    If frmMain.MousePointer <> 2 Then Exit Sub 'Parcheo porque a veces tira el hechizo sin tener el cursor (NicoNZ)
                    
                    frmMain.MousePointer = vbDefault
                    Call WriteWorkLeftClick(tX, tY, UsingSkill)
                    UsingSkill = 0
                End If
            Else
                Call AbrirMenuViewPort
            End If
        ElseIf (MouseShift And 1) = 1 Then
            If Not CustomKeys.KeyAssigned(KeyCodeConstants.vbKeyShift) Then
                If MouseBoton = vbLeftButton Then
                    Call WriteWarpChar("YO", UserMap, tX, tY)
                End If
            End If
        End If
    End If
End Sub
Private Sub Form_DblClick()
'**************************************************************
'Author: Unknown
'Last Modify Date: 12/27/2007
'12/28/2007: ByVal - Chequea que la ventana de comercio y boveda no este abierta al hacer doble clic a un comerciante, sobrecarga la lista de items.
'**************************************************************
    If Not frmForo.Visible And Not Comerciando Then 'frmComerciar.Visible And Not frmBancoObj.Visible Then
        Call WriteDoubleClick(tX, tY)
    End If
End Sub

Private Sub Form_Load()
    Call Make_Transparent_Richtext(RecTxt.hwnd)
    
    Me.Picture = General_Load_Picture_From_Resource("42.gif")
    ExpShp.Picture = General_Load_Picture_From_Resource("104.gif")
    Hpshp.Picture = General_Load_Picture_From_Resource("105.gif")
    MANShp.Picture = General_Load_Picture_From_Resource("106.gif")
    COMIDAsp.Picture = General_Load_Picture_From_Resource("107.gif")
    AGUAsp.Picture = General_Load_Picture_From_Resource("108.gif")
    STAShp.Picture = General_Load_Picture_From_Resource("109.gif")
    BarraAgilidad.Picture = General_Load_Picture_From_Resource("110.gif")
    BarraFuerza.Picture = General_Load_Picture_From_Resource("111.gif")
    PicResu.Picture = General_Load_Picture_From_Resource("112.gif")
    
    Me.Caption = Form_Caption
    
    Me.Left = 0
    Me.Top = 0
   
    Cargar_List

End Sub

Private Sub LComm_Click()
    If Opciones.AutoComandos = False Then
        SendTxt.Text = LComm.Text
    End If
End Sub
 
Private Sub LComm_DblClick()
    If Opciones.AutoComandos = False Then
        SendTxt.Text = LComm.Text
        LComm.Visible = False
    End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseX = X - MainViewPic.Left
    MouseY = Y - MainViewPic.Top
    
    'Trim to fit screen
    If MouseX < 0 Then
        MouseX = 0
    ElseIf MouseX > MainViewPic.Width Then
        MouseX = MainViewPic.Width
    End If
    
    'Trim to fit screen
    If MouseY < 0 Then
        MouseY = 0
    ElseIf MouseY > MainViewPic.Height Then
        MouseY = MainViewPic.Height
    End If
    
    
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
    Call General_Set_Wav(SND_CLICK)

    Select Case Index
        Case 0
            Call frmOpciones.Show(vbModeless, frmMain)
            
        Case 1
            LlegaronAtrib = False
            LlegaronSkills = False
            LlegoFama = False
            Call WriteRequestAtributes
            Call WriteRequestSkills
            Call WriteRequestMiniStats
            Call WriteRequestFame
            Call FlushBuffer
            Do While Not LlegaronSkills Or Not LlegaronAtrib Or Not LlegoFama
                DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
            Loop
            Alocados = SkillPoints
            frmEstadisticas.Puntos.Caption = SkillPoints
            frmEstadisticas.Iniciar_Labels
            frmEstadisticas.Show , frmMain
            LlegaronAtrib = False
            LlegaronSkills = False
            LlegoFama = False
        
        Case 2
            If frmGuildLeader.Visible Then Unload frmGuildLeader
            
            Call WriteRequestGuildLeaderInfo
    End Select
End Sub

Private Sub Image3_Click(Index As Integer)
    Select Case Index
        Case 0
            Inventario.SelectGold
            If UserGLD > 0 Then
                frmCantidad.Show , frmMain
            End If
    End Select
End Sub

Private Sub Label4_Click()
    Call General_Set_Wav(SND_CLICK)

    InvEqu.Picture = General_Load_Picture_From_Resource("57.gif")
    ' Activo controles de inventario
    PicInv.Visible = True
    
    ' Desactivo controles de hechizo
    hlst.Visible = False
    cmdINFO.Visible = False
    CmdLanzar.Visible = False
    
    cmdMoverHechi(0).Visible = False
    cmdMoverHechi(1).Visible = False
    
    ItemInfo.Visible = True
    
    Inventario.UpdateInventory
        
End Sub

Private Sub Label7_Click()
    Call General_Set_Wav(SND_CLICK)

    InvEqu.Picture = General_Load_Picture_From_Resource("56.gif")
    '%%%%%%OCULTAMOS EL INV&&&&&&&&&&&&
    'DespInv(0).Visible = False
    'DespInv(1).Visible = False
    PicInv.Visible = False
    hlst.Visible = True
    cmdINFO.Visible = True
    CmdLanzar.Visible = True
    
    cmdMoverHechi(0).Visible = True
    cmdMoverHechi(1).Visible = True
    
    cmdMoverHechi(0).Enabled = True
    cmdMoverHechi(1).Enabled = True
    
    ItemInfo.Visible = False
End Sub
Private Sub picInv_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseBoton = Button
End Sub
Private Sub picInv_DblClick()
    If frmCarp.Visible Or frmHerrero.Visible Then Exit Sub
    
    If Not MainTimer.Check(TimersIndex.UseItemWithDblClick) Then Exit Sub
    
    If MainTimer.Check(TimersIndex.UseItemWithU) Then Call UsarItem
    Call EquiparItem
End Sub

Private Sub picInv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call General_Set_Wav(SND_CLICK)
End Sub

Private Sub RecTxt_Change()
On Error Resume Next  'el .SetFocus causaba errores al salir y volver a entrar
    If Not Multimod.IsAppActive() Then Exit Sub
    
    If SendTxt.Visible Then
        SendTxt.SetFocus
    ElseIf (Not Comerciando) And (Not frmCantidad.Visible) And _
        (Not frmMSG.Visible) And (Not frmForo.Visible) And _
        (Not frmEstadisticas.Visible) Then
         
        If PicInv.Visible Then
            PicInv.SetFocus
        ElseIf hlst.Visible Then
            hlst.SetFocus
        End If
    End If
End Sub

Private Sub RecTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If PicInv.Visible Then
        PicInv.SetFocus
    Else
        hlst.SetFocus
    End If
End Sub

Private Sub SendTxt_Change()
'**************************************************************
'Author: Unknown
'Last Modify Date: 3/06/2006
'3/06/2006: Maraxus - impedí se inserten caractéres no imprimibles
'**************************************************************
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
        If MainTimer.Check(TimersIndex.UseItemWithU) Then Call UsarItem
    Case 3 'equipar
        Call EquiparItem
    End Select
    
Case 1 'Menu del ViewPort del engine
    Select Case Sel
    Case 0 'Nombre
        Call WriteLeftClick(tX, tY)
        
    Case 1 'Comerciar
        Call WriteLeftClick(tX, tY)
        Call WriteCommerceStart
    End Select
End Select
End Sub


'
' -------------------
'    W I N S O C K
' -------------------
'

Private Sub Winsock1_Close()
    Dim i As Long
    
    Debug.Print "WInsock Close"
    
    Connected = False
    
    If Winsock1.State <> sckClosed Then _
        Winsock1.Close
    
    frmConnect.MousePointer = vbNormal
    
    If Not frmCrearPersonaje.Visible Then
        General_Set_Song 2, True
        Call Relog
    End If
    
    Do While i < Forms.count - 1
        i = i + 1
        
        If Forms(i).name <> Me.name And Forms(i).name <> frmConnect.name And Forms(i).name <> frmCrearPersonaje.name And Forms(i).name <> frmMain.name And Forms(i).name <> frmCuenta.name Then
            Unload Forms(i)
        End If
    Loop
    On Local Error GoTo 0
    
    Call SetMusicInfo("", "", "", "Games", , False)
    frmMain.Visible = False

    pausa = False
    UserMeditar = False

    UserClase = 0
    UserSexo = 0
    UserRaza = 0
    UserEmail = ""
    UserEquitando = False
    'Reseteo el Speed
    Call SpeedCaballo
    bTechoAB = 255
    
    For i = 1 To NUMSKILLS
        UserSkills(i) = 0
    Next i

    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i

    SkillPoints = 0
    Alocados = 0
End Sub

Private Sub Winsock1_Connect()
    Debug.Print "Winsock Connect"
   
    'Clean input and output buffers
    Call incomingData.ReadASCIIStringFixed(incomingData.Length)
    Call outgoingData.ReadASCIIStringFixed(outgoingData.Length)
 
   
    Select Case EstadoLogin
        Case E_MODO.Dados
            frmCrearPersonaje.Show vbModal
        Case Else
            Call Login
    End Select
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim RD As String
    Dim Data() As Byte
    
    Winsock1.GetData RD
    
    Data = StrConv(RD, vbFromUnicode)
    
    'Set data in the buffer
    Call incomingData.WriteBlock(Data)
    
    'Send buffer to Handle data
    Call HandleIncomingData
End Sub

Private Sub Winsock1_Error(ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    '*********************************************
    'Handle socket errors
    '*********************************************
    Dim i As Long
    
    frmMensaje.msg.Caption = Description
    frmMensaje.Show vbModal
    
    frmConnect.MousePointer = 1

    If Winsock1.State <> sckClosed Then _
        Winsock1.Close

    If Not frmCrearPersonaje.Visible Then
        If frmConnect.Visible Then Exit Sub
        General_Set_Song 2, True
        Call Audio.StopWave
        
        If frmCuenta.Visible Then
            Unload frmCuenta
        ElseIf frmMain.Visible Then
            frmMain.Visible = False
        End If
        
        Call SetMusicInfo("", "", "", "Games", , False)
        pausa = False
        
        UserClase = 0
        UserSexo = 0
        UserRaza = 0
        UserEmail = ""
        bTechoAB = 255
        
        'Reset global vars
        IScombate = False
        UserDescansar = False
        UserParalizado = False
        pausa = False
        UserCiego = False
        UserMeditar = False
        UserNavegando = False
        UserEquitando = False
        'Reseteo el Speed
        Call SpeedCaballo
        bFogata = False
        SkillPoints = 0
            
            
        For i = 1 To NUMSKILLS
            UserSkills(i) = 0
        Next i
        
        For i = 1 To NUMATRIBUTOS
            UserAtributos(i) = 0
        Next i
        
        SkillPoints = 0
        Alocados = 0
        
        frmConnect.Show
    Else
        frmCrearPersonaje.MousePointer = 0
    End If
End Sub

Private Function InGameArea() As Boolean
'***************************************************
'Author: NicoNZ
'Last Modification: 04/07/08
'Checks if last click was performed within or outside the game area.
'***************************************************
    If clicX < MainViewPic.Left Or clicX > MainViewPic.Left + MainViewPic.Width Then Exit Function
    If clicY < MainViewPic.Top Or clicY > MainViewPic.Top + MainViewPic.Height Then Exit Function
    
    InGameArea = True
End Function
Public Sub Client_Screenshot(ByVal hDC As Long, ByVal Width As Long, ByVal Height As Long)
On Error GoTo ErrorHandler

Dim i As Long
Dim Index As Long
i = 1

Set m_Jpeg = New clsJpeg

'80 Quality
m_Jpeg.Quality = 100

'Sample the cImage by hDC
m_Jpeg.SampleHDC hDC, Width, Height

m_FileName = App.Path & "\Fotos\WinterAO_Foto"

If Dir(App.Path & "\Fotos", vbDirectory) = vbNullString Then
    MkDir (App.Path & "\Fotos")
End If

Do While Dir(m_FileName & Trim(str(i)) & ".jpg") <> vbNullString
    i = i + 1
    DoEvents
Loop

Index = i

m_Jpeg.Comment = "Character: " & PJName & " - " & Format(Date, "dd/mm/yyyy") & " - " & Format(Time, "hh:mm AM/PM")

'Save the JPG file
m_Jpeg.SaveFile m_FileName & Trim(str(Index)) & ".jpg"

Call AddtoRichTextBox(frmMain.RecTxt, "La foto fue guardada en " & m_FileName & Trim(str(Index)) & ".jpg", 65, 190, 156, False, True, False)

Set m_Jpeg = Nothing

Exit Sub

ErrorHandler:
    Call AddtoRichTextBox(frmMain.RecTxt, "¡Error al tomar la foto!", 65, 190, 156, False, True, False)

End Sub
