VERSION 5.00
Begin VB.Form frmCanjes 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2280
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   ScaleHeight     =   2280
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Otros 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   5
      Left            =   1080
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   36
      Top             =   1200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox Otros 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   4
      Left            =   480
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   35
      Top             =   1200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox Otros 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   3
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   34
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox Otros 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   2
      Left            =   1680
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   33
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox Otros 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   1
      Left            =   1050
      Picture         =   "frmCanjes.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   32
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox Otros 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   0
      Left            =   420
      Picture         =   "frmCanjes.frx":0C44
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   31
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox Armas 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   5
      Left            =   1080
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   30
      Top             =   1200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox Armas 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   4
      Left            =   1800
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   29
      Top             =   1200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox Armas 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   3
      Left            =   2280
      Picture         =   "frmCanjes.frx":1486
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   28
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox Armas 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   2
      Left            =   1680
      Picture         =   "frmCanjes.frx":1CC8
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   27
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox Armas 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   1
      Left            =   1050
      Picture         =   "frmCanjes.frx":290A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   26
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox Armas 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   0
      Left            =   450
      Picture         =   "frmCanjes.frx":314C
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   25
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox Armaduras 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   5
      Left            =   1080
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   24
      Top             =   1200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox Armaduras 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   4
      Left            =   480
      Picture         =   "frmCanjes.frx":3D8E
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   23
      Top             =   1200
      Width           =   480
   End
   Begin VB.PictureBox Armaduras 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   3
      Left            =   2260
      Picture         =   "frmCanjes.frx":49D0
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   22
      Top             =   600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox Armaduras 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   2
      Left            =   1680
      Picture         =   "frmCanjes.frx":52B6
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   21
      Top             =   600
      Width           =   480
   End
   Begin VB.PictureBox Armaduras 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   1
      Left            =   1040
      Picture         =   "frmCanjes.frx":5AFA
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   20
      Top             =   600
      Width           =   480
   End
   Begin VB.PictureBox Armaduras 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   0
      Left            =   420
      Picture         =   "frmCanjes.frx":5F70
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   19
      Top             =   600
      Width           =   480
   End
   Begin VB.PictureBox Cascos 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   5
      Left            =   1080
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   17
      Top             =   1200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox Cascos 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   4
      Left            =   470
      Picture         =   "frmCanjes.frx":63E6
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   16
      Top             =   1200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox Cascos 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   3
      Left            =   2270
      Picture         =   "frmCanjes.frx":6C28
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   15
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox Cascos 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   2
      Left            =   1640
      Picture         =   "frmCanjes.frx":786A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   14
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox Cascos 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   1
      Left            =   1040
      Picture         =   "frmCanjes.frx":806E
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   13
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox Cascos 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   0
      Left            =   420
      Picture         =   "frmCanjes.frx":8CB0
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   12
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox Monturas 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   5
      Left            =   1080
      Picture         =   "frmCanjes.frx":98F4
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   11
      Top             =   1200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox Monturas 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   4
      Left            =   460
      Picture         =   "frmCanjes.frx":A938
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   10
      Top             =   1200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox Monturas 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   3
      Left            =   2260
      Picture         =   "frmCanjes.frx":B57C
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   9
      Top             =   580
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox Monturas 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   2
      Left            =   1650
      Picture         =   "frmCanjes.frx":C1C0
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   8
      Top             =   580
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox Monturas 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   0
      Left            =   420
      Picture         =   "frmCanjes.frx":CE04
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox Monturas 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   1
      Left            =   1040
      Picture         =   "frmCanjes.frx":DA46
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image5 
      Height          =   255
      Left            =   4440
      Top             =   240
      Width           =   855
   End
   Begin VB.Image Image4 
      Height          =   255
      Left            =   2520
      Top             =   240
      Width           =   735
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   360
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Precio 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Precio: 0"
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
      Left            =   2610
      TabIndex        =   18
      Top             =   1905
      Width           =   795
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   1680
      Top             =   240
      Width           =   735
   End
   Begin VB.Image Command1 
      Height          =   255
      Left            =   240
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Image Cerrar 
      Height          =   255
      Left            =   4440
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Informacion 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
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
      Height          =   495
      Left            =   3000
      TabIndex        =   7
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Información:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Puntos disponibles:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   5
      Top             =   600
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   3280
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Nombre 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
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
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   3
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Puntos 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   4680
      TabIndex        =   1
      Top             =   610
      Width           =   975
   End
End
Attribute VB_Name = "frmCanjes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Seleccionado As Byte

Private Sub Armaduras_Click(Index As Integer)
    Select Case Index
        Case 0
            Seleccionado = 11
            Nombre.Caption = "Armadura Logouth"
            Informacion.Caption = "Def: 68/63, Nivel:50"
            Precio.Caption = "Precio: 85pts"
        Case 1
            Seleccionado = 12
            Nombre.Caption = "Armadura Logouth (Bajos)"
            Informacion.Caption = "Def: 68/63, Nivel:50"
            Precio.Caption = "Precio: 85pts"
        Case 2
            Seleccionado = 13
            Nombre.Caption = "Túnica de Dragon (H/M)"
            Informacion.Caption = "Def: 68/63, Nivel:50"
            Precio.Caption = "Precio: 85pts"
        Case 3
            Seleccionado = 14
            Nombre.Caption = "Túnica de Dragon (Bajos)(H/M)"
            Informacion.Caption = "Def: 68/63, Nivel:50"
            Precio.Caption = "Precio: 85pts"
        Case 4
            Seleccionado = 15
            Nombre.Caption = "Túnica de Akatsuki(H/M)"
            Informacion.Caption = "Def: 25/15"
            Precio.Caption = "Precio: 30pts"
      End Select
            
End Sub

Private Sub Armas_Click(Index As Integer)
    Select Case Index
        Case 0
            Seleccionado = 16
            Nombre.Caption = "Arco Argentum"
            Informacion.Caption = "Def: 15/10, Nivel:50"
            Precio.Caption = "Precio: 85pts"
        Case 1
            Seleccionado = 17
            Nombre.Caption = "Espada Argentum"
            Informacion.Caption = "Def: 29/25, Nivel:50"
            Precio.Caption = "Precio: 85pts"
        Case 2
            Seleccionado = 18
            Nombre.Caption = "Vara del Mago Legendario"
            Informacion.Caption = "Def: 16/7, Nivel:50"
            Precio.Caption = "Precio: 85pts"
        Case 3
            Seleccionado = 19
            Nombre.Caption = "Daga Demoniaca"
            Informacion.Caption = "Def: 15/11, Nivel:45"
            Precio.Caption = "Precio: 75pts"
    End Select
End Sub

Private Sub Cerrar_Click()
Unload Me
Comerciando = False
End Sub
Private Sub Command1_Click()
If Seleccionado = 0 Then
    MsgBox "Selecciona un objeto", vbInformation
    Exit Sub
Else
    Call WriteCanjes(Seleccionado)
End If
End Sub

Private Sub Form_Load()
    Me.Picture = General_Load_Picture_From_Resource("33.gif")
End Sub

Private Sub Image1_Click()
    Dim i As Byte
    
    For i = 0 To 5
        Armaduras(i).Visible = False
        Cascos(i).Visible = False
        Armas(i).Visible = False
        Monturas(i).Visible = True
    Next i
    
    Me.Picture = General_Load_Picture_From_Resource("36.gif")
End Sub

Private Sub Image2_Click()
    Dim i As Byte
    
    For i = 0 To 5
        Armaduras(i).Visible = False
        Cascos(i).Visible = True
        Armas(i).Visible = False
        Monturas(i).Visible = False
    Next i
    
    Me.Picture = General_Load_Picture_From_Resource("34.gif")
End Sub

Private Sub Cascos_Click(Index As Integer)
    Select Case Index
        Case 0
            Seleccionado = 6
            Nombre.Caption = "Sombrero de Mago (+RM)"
            Informacion.Caption = "Def: 28/24, Nivel:50"
            Precio.Caption = "Precio: 65pts"
        Case 1
            Seleccionado = 7
            Nombre.Caption = "Casco Alethril"
            Informacion.Caption = "Def: 50/48, Nivel:50"
            Precio.Caption = "Precio: 65pts"
        Case 2
            Seleccionado = 8
            Nombre.Caption = "Casco de Astas"
            Informacion.Caption = "Def: 48/45, Nivel:48"
            Precio.Caption = "Precio: 60pts"
        Case 3
            Seleccionado = 9
            Nombre.Caption = "Casco Alado"
            Informacion.Caption = "Def: 45/43, Nivel:45"
            Precio.Caption = "Precio: 55pts"
        Case 4
            Seleccionado = 10
            Nombre.Caption = "Gafas"
            Informacion.Caption = "Def: 1/1, Nivel:1. ¡Presume de tus gafas!"
            Precio.Caption = "Precio: 35pts"
    End Select
End Sub

Private Sub Image3_Click()
    Dim i As Byte
    
    For i = 0 To 5
        Armaduras(i).Visible = True
        Cascos(i).Visible = False
        Armas(i).Visible = False
        Monturas(i).Visible = False
        Otros(i).Visible = False
    Next i

    Me.Picture = General_Load_Picture_From_Resource("33.gif")
End Sub

Private Sub Image4_Click()
    Dim i As Byte
    
    For i = 0 To 5
        Armaduras(i).Visible = False
        Cascos(i).Visible = False
        Armas(i).Visible = True
        Monturas(i).Visible = False
        Otros(i).Visible = False
    Next i

    Me.Picture = General_Load_Picture_From_Resource("35.gif")
End Sub

Private Sub Image5_Click()
    Dim i As Byte
    
    For i = 0 To 5
        Armaduras(i).Visible = False
        Cascos(i).Visible = False
        Armas(i).Visible = False
        Monturas(i).Visible = False
        Otros(i).Visible = True
    Next i

    Me.Picture = General_Load_Picture_From_Resource("37.gif")
End Sub

Private Sub Monturas_Click(Index As Integer)
'********************************************
'MONTURAS
'********************************************
    Select Case Index
    
        Case 0
            Seleccionado = 1
            Nombre.Caption = "Montura Preclitus"
            Informacion.Caption = "Skills:100, Nivel:50. Aumenta velocidad, golpe y defensa."
            Precio.Caption = "Precio: 75 pts"
        Case 1
            Seleccionado = 2
            Nombre.Caption = "Montura Preclitus Azul"
            Informacion.Caption = "Skills:100, Nivel:50. Aumenta velocidad, golpe y defensa."
            Precio.Caption = "Precio: 75pts"
        Case 2
            Seleccionado = 3
            Nombre.Caption = "Montura de Dragón Negro"
            Informacion.Caption = "Skills:100, Nivel:50. Aumenta velocidad, golpe y defensa."
            Precio.Caption = "Precio: 100pts"
        Case 3
            Seleccionado = 4
            Nombre.Caption = "Montura de Dragón Dorado"
            Informacion.Caption = "Skills:100, Nivel:50. Aumenta velocidad, golpe y defensa."
            Precio.Caption = "Precio: 100pts"
        Case 4
            Seleccionado = 5
            Nombre.Caption = "Montura de Dragón Rojo"
            Informacion.Caption = "Skills:100, Nivel:50. Aumenta velocidad, golpe y defensa."
            Precio.Caption = "Precio: 100pts"
        Case 5
            Seleccionado = 6
            Nombre.Caption = "Montura de Buey"
            Informacion.Caption = "Skills:100, Nivel:50. Aumenta velocidad, golpe y defensa."
            Precio.Caption = "Precio: 85pts"
    End Select
End Sub

Private Sub Otros_Click(Index As Integer)
    Select Case Index
    
        Case 0
            Seleccionado = 20
            Nombre.Caption = "Pendiente del sacrificio"
            Informacion.Caption = "Al morir solo perderas el pendiente."
            Precio.Caption = "Precio: 30 pts"
        Case 1
            Seleccionado = 21
            Nombre.Caption = "Anillo de la hermandad"
            Informacion.Caption = "Requerido para fundar clan."
            Precio.Caption = "Precio: 100 pts"
    End Select
End Sub
