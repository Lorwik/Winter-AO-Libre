VERSION 5.00
Begin VB.Form frmSubasta 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmSubasta.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "frmSubasta.frx":0CCA
   ScaleHeight     =   5205
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TextBox2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   240
      Left            =   3480
      TabIndex        =   9
      Text            =   "1"
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox TextBox1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   240
      Left            =   3480
      TabIndex        =   8
      Text            =   "1"
      Top             =   2640
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   2670
      Index           =   1
      Left            =   720
      TabIndex        =   1
      Top             =   1200
      Width           =   2325
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   500
      Left            =   3500
      ScaleHeight     =   465
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   1090
      Width           =   505
   End
   Begin VB.Image Image1 
      Height          =   690
      Index           =   1
      Left            =   3600
      MouseIcon       =   "frmSubasta.frx":1BFE6
      MousePointer    =   99  'Custom
      Picture         =   "frmSubasta.frx":1CCB0
      Tag             =   "1"
      Top             =   3600
      Width           =   2610
   End
   Begin VB.Image Image2 
      Height          =   225
      Left            =   6240
      MouseIcon       =   "frmSubasta.frx":1E630
      MousePointer    =   99  'Custom
      Picture         =   "frmSubasta.frx":1F2FA
      Top             =   240
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   810
      TabIndex        =   7
      Top             =   120
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   1
      Left            =   1485
      TabIndex        =   6
      Top             =   420
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "None"
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
      Index           =   2
      Left            =   3480
      TabIndex        =   5
      Top             =   2100
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   4
      Left            =   3435
      TabIndex        =   4
      Top             =   1140
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   3
      Left            =   3435
      TabIndex        =   3
      Top             =   1485
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1950
      TabIndex        =   2
      Top             =   6420
      Width           =   645
   End
End
Attribute VB_Name = "frmSubasta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'<-------------------------NUEVO-------------------------->
Public LastIndex1 As Integer
Public LastIndex2 As Integer

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
        Select Case Index
        Case 1
        If Image1(1).Tag = 1 Then
                Image1(1).Picture = General_Load_Picture_From_Resource("subaoff.gif")
                Image1(1).Tag = 0
        End If
        End Select
                End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If Image1(1).Tag = 0 Then
                Image1(1).Picture = General_Load_Picture_From_Resource("suba.gif")
                Image1(1).Tag = 1
        End If
End Sub
Private Sub Form_Load()
Image1(1).Picture = General_Load_Picture_From_Resource("suba.gif")
End Sub

Private Sub Image1_Click(Index As Integer)

Call Audio.PlayWave(SND_CLICK)

If List1(Index).List(List1(Index).listIndex) = "Nada" Or _
   List1(Index).listIndex < 0 Then Exit Sub

Select Case Index
   Case 1
   Image1(1).Picture = General_Load_Picture_From_Resource("subaoff.gif")
        LastIndex2 = List1(1).listIndex
        If Not Inventario.Equipped(List1(1).listIndex + 1) Then
            SendData ("SUBA" & "," & List1(1).listIndex + 1 & "," & TextBox1.Text & "," & TextBox2.Text)
        Else
            AddtoRichTextBox frmMain.RecTxt, "No podes vender el item porque lo estas usando.", 2, 51, 223, 1, 1
            Exit Sub
        End If

End Select

List1(1).Clear

frmMain.SetFocus
Unload Me

NPCInvDim = 0
End Sub

Private Sub Image2_Click()
SendData ("FINSUB")
End Sub

Private Sub list1_Click(Index As Integer)
Dim SR As RECT, DR As RECT

SR.Left = 0
SR.Top = 0
SR.Right = 32
SR.Bottom = 32

DR.Left = 0
DR.Top = 0
DR.Right = 32
DR.Bottom = 32

Select Case Index
    Case 1
        Call DrawGrhtoHdc(Picture1.hwnd, Picture1.hDC, Inventario.GrhIndex(List1(1).listIndex + 1), SR, DR)
End Select

Picture1.Refresh

End Sub

Private Sub TextBox1_Change()
If Val(TextBox1.Text) < 1 Then
        TextBox1.Text = 1
    End If
    
    If Val(TextBox1.Text) > MAX_INVENTORY_OBJS Then
        TextBox1.Text = 1
    End If
End Sub


