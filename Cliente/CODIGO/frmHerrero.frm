VERSION 5.00
Begin VB.Form frmHerrero 
   BorderStyle     =   0  'None
   Caption         =   "Herrero"
   ClientHeight    =   4425
   ClientLeft      =   0
   ClientTop       =   -45
   ClientWidth     =   5625
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Cantidad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Height          =   240
      Left            =   2470
      MaxLength       =   6
      TabIndex        =   2
      Text            =   "1"
      ToolTipText     =   "Ingrese la cantidad total de items a construir."
      Top             =   3930
      Width           =   680
   End
   Begin VB.ListBox lstArmas 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   3345
      Left            =   690
      TabIndex        =   0
      Top             =   480
      Width           =   4200
   End
   Begin VB.ListBox lstArmaduras 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   3345
      Left            =   690
      TabIndex        =   1
      Top             =   480
      Width           =   4200
   End
   Begin VB.Image Command2 
      Height          =   255
      Left            =   1560
      Top             =   240
      Width           =   1095
   End
   Begin VB.Image Command1 
      Height          =   255
      Left            =   720
      Top             =   240
      Width           =   780
   End
   Begin VB.Image Command4 
      Height          =   255
      Left            =   960
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Image Command3 
      Height          =   255
      Left            =   3240
      Top             =   3960
      Width           =   1455
   End
End
Attribute VB_Name = "frmHerrero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Call General_Set_Wav(SND_CLICK)
lstArmaduras.Visible = False
lstArmas.Visible = True
Me.Picture = General_Load_Picture_From_Resource("69.gif")
End Sub

Private Sub Command2_Click()
Call General_Set_Wav(SND_CLICK)
lstArmaduras.Visible = True
lstArmas.Visible = False
Me.Picture = General_Load_Picture_From_Resource("68.gif")
End Sub

Private Sub Command3_Click()
On Error Resume Next
Call General_Set_Wav(SND_CLICK)

    If lstArmas.Visible Then
        Call WriteCraftBlacksmith(ArmasHerrero(lstArmas.ListIndex + 1), cantidad.Text)
    Else
        Call WriteCraftBlacksmith(ArmadurasHerrero(lstArmaduras.ListIndex + 1), cantidad.Text)
        
    End If
    
    Unload Me
End Sub

Private Sub Command4_Click()
Call General_Set_Wav(SND_CLICK)
Unload Me
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = vbLeftButton) Then Call Auto_Drag(Me.hwnd)
End Sub
Private Sub Form_Load()
Me.Picture = General_Load_Picture_From_Resource("68.gif")
End Sub
