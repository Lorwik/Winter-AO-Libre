VERSION 5.00
Begin VB.Form frmBorrarPj 
   BorderStyle     =   0  'None
   ClientHeight    =   3750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox CodeKey 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1350
      MaxLength       =   6
      TabIndex        =   2
      Top             =   2950
      Width           =   2685
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2020
      TabIndex        =   1
      Top             =   1500
      Width           =   1335
   End
   Begin VB.Image Command2 
      Height          =   255
      Left            =   360
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Image Command1 
      Height          =   255
      Left            =   3720
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label lblCode 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   2460
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2115
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmBorrarPj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = vbLeftButton) Then Call Auto_Drag(Me.hWnd)
End Sub
Private Sub Command1_Click()
Call General_Set_Wav(SND_CLICK)
        If Not Text1.Text = "BORRAR" Then
            MsgBox "Debe de escribir BORRAR Para borrar el personaje"
            Exit Sub
        End If
        
        If lblCode.Caption <> CodeKey.Text Then
            MsgBox "El Codigo ingresado es Invalido.", vbCritical
            lblCode.Caption = GenerateKey
            Exit Sub
        End If
        
        If MsgBox("Al borrar un personaje de su cuenta perderá todo lo que hay en él." & vbCrLf & "¿Seguro que desea borrar el personaje " & PJName & "?", vbInformation + vbYesNo, "Eliminar Personaje de la cuenta.") = vbYes Then
            
            EstadoLogin = BorrarPJ
            
            frmMain.Winsock1.Connect CurServerIp, CurServerPort
            
                Unload Me
                DoEvents
                MsgBox "Relogee para ver cambios.", vbInformation
            Exit Sub
        Else
            Unload Me
            Exit Sub
        End If
        
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Picture = General_Load_Picture_From_Resource("62.gif")
Label2.Caption = PJName
lblCode.Caption = GenerateKey
End Sub

Private Sub Text1_Change()
Text1.Text = LTrim(Text1.Text)
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
