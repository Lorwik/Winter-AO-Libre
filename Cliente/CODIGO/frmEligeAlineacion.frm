VERSION 5.00
Begin VB.Form frmEligeAlineacion 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5895
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6930
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   6930
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Elige una Alienación"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   240
      TabIndex        =   11
      Top             =   0
      Width           =   6255
   End
   Begin VB.Label lblSalir 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5400
      TabIndex        =   10
      Top             =   5460
      Width           =   915
   End
   Begin VB.Label lblDescripcion 
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmEligeAlineacion.frx":0000
      ForeColor       =   &H00FFFFFF&
      Height          =   645
      Index           =   4
      Left            =   1095
      TabIndex        =   9
      Top             =   4740
      Width           =   5505
   End
   Begin VB.Label lblDescripcion 
      BackColor       =   &H00000080&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmEligeAlineacion.frx":00D5
      ForeColor       =   &H00FFFFFF&
      Height          =   645
      Index           =   3
      Left            =   1095
      TabIndex        =   8
      Top             =   3840
      Width           =   5505
   End
   Begin VB.Label lblDescripcion 
      BackColor       =   &H00400040&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmEligeAlineacion.frx":01B1
      ForeColor       =   &H00FFFFFF&
      Height          =   645
      Index           =   2
      Left            =   1095
      TabIndex        =   7
      Top             =   2895
      Width           =   5505
   End
   Begin VB.Label lblDescripcion 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmEligeAlineacion.frx":025D
      ForeColor       =   &H00FFFFFF&
      Height          =   645
      Index           =   1
      Left            =   1095
      TabIndex        =   6
      Top             =   1950
      Width           =   5505
   End
   Begin VB.Label lblDescripcion 
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmEligeAlineacion.frx":0326
      ForeColor       =   &H00FFFFFF&
      Height          =   825
      Index           =   0
      Left            =   1095
      TabIndex        =   5
      Top             =   870
      Width           =   5505
   End
   Begin VB.Label lblNombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Alineación del mal"
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
      Index           =   4
      Left            =   1005
      TabIndex        =   4
      Top             =   4515
      Width           =   1680
   End
   Begin VB.Label lblNombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Alineación criminal"
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
      Index           =   3
      Left            =   1005
      TabIndex        =   3
      Top             =   3615
      Width           =   1680
   End
   Begin VB.Label lblNombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Alineación neutral"
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
      Index           =   2
      Left            =   1005
      TabIndex        =   2
      Top             =   2670
      Width           =   1635
   End
   Begin VB.Label lblNombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Alineación legal"
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
      Index           =   1
      Left            =   1005
      TabIndex        =   1
      Top             =   1725
      Width           =   1455
   End
   Begin VB.Label lblNombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Alineación Real"
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
      Index           =   0
      Left            =   1005
      TabIndex        =   0
      Top             =   645
      Width           =   1455
   End
End
Attribute VB_Name = "frmEligeAlineacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Byte

    For i = 0 To 4
        lblDescripcion(i).BorderStyle = 0
        lblDescripcion(i).BackStyle = 0
    Next i
    
End Sub
Private Sub lblDescripcion_Click(Index As Integer)
Dim S As String
    
    Select Case Index
        Case 0
            S = "armada"
        Case 1
            S = "legal"
        Case 2
            S = "neutro"
        Case 3
            S = "criminal"
        Case 4
            S = "mal"
    End Select
    
    S = "/fundarclan " & S
    Call SendData(S)
    Unload Me
End Sub
Private Sub lblDescripcion_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblDescripcion(Index).BorderStyle = 1
    lblDescripcion(Index).BackStyle = 1
    Select Case Index
        Case 0
            lblDescripcion(Index).BackColor = &H400000
        Case 1
            lblDescripcion(Index).BackColor = &H800000
        Case 2
            lblDescripcion(Index).BackColor = 4194368
        Case 3
            lblDescripcion(Index).BackColor = &H80&
        Case 4
            lblDescripcion(Index).BackColor = &H40&
    End Select
End Sub
Private Sub lblNombre_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Byte

    For i = 0 To 4
        lblDescripcion(i).BorderStyle = 0
        lblDescripcion(i).BackStyle = 0
    Next i
    

End Sub
Private Sub lblSalir_Click()
    Unload Me
End Sub
