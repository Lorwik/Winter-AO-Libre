VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCargando 
   AutoRedraw      =   -1  'True
   BackColor       =   &H000080FF&
   BorderStyle     =   0  'None
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4440
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCargando.frx":0000
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   296
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox Status 
      Height          =   720
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   1080
      Width           =   4440
      _ExtentX        =   7832
      _ExtentY        =   1270
      _Version        =   393217
      BackColor       =   14737632
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmCargando.frx":0C7C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox LOGO 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   0
      Left            =   9645
      ScaleHeight     =   0
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   0
      TabIndex        =   0
      Top             =   7260
      Width           =   0
   End
End
Attribute VB_Name = "frmCargando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Const LW_KEY = &H1
Const G_E = (-20)
Const W_E = &H80000

Private Sub Form_Load()
Skin Me, vbRed
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'para mover el form de cualquier parte
ReleaseCapture
SendMessage hWnd, 161, 2, 0
End Sub
Sub Skin(Frm As Form, Color As Long)
Frm.BackColor = Color
Dim Ret As Long
Ret = GetWindowLong(Frm.hWnd, G_E)
Ret = Ret Or W_E
SetWindowLong Frm.hWnd, G_E, Ret
SetLayeredWindowAttributes Frm.hWnd, Color, 0, LW_KEY
End Sub
