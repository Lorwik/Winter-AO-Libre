VERSION 5.00
Begin VB.Form frmCrearPersonaje 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   599
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Cabeza 
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":0000
      Left            =   12000
      List            =   "frmCrearPersonaje.frx":0002
      TabIndex        =   12
      Top             =   8880
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.PictureBox PlayerView 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   10095
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   11
      Top             =   6105
      Width           =   240
   End
   Begin VB.ComboBox lstProfesion 
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
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":0004
      Left            =   9015
      List            =   "frmCrearPersonaje.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   5100
      Width           =   2340
   End
   Begin VB.ComboBox lstGenero 
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
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":0008
      Left            =   9015
      List            =   "frmCrearPersonaje.frx":0012
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   4170
      Width           =   2340
   End
   Begin VB.ComboBox lstRaza 
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
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":0025
      Left            =   9015
      List            =   "frmCrearPersonaje.frx":0027
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3360
      Width           =   2340
   End
   Begin VB.TextBox txtNombre 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   6060
      TabIndex        =   0
      Top             =   1680
      Width           =   4065
   End
   Begin VB.Label ModCarisma 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3210
      TabIndex        =   18
      Top             =   5610
      Width           =   90
   End
   Begin VB.Label ModConstitucion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3210
      TabIndex        =   17
      Top             =   5220
      Width           =   90
   End
   Begin VB.Label ModInteligencia 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3210
      TabIndex        =   16
      Top             =   4815
      Width           =   90
   End
   Begin VB.Label ModAgilidad 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3210
      TabIndex        =   15
      Top             =   4410
      Width           =   90
   End
   Begin VB.Label ModFuerza 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3210
      TabIndex        =   14
      Top             =   4020
      Width           =   90
   End
   Begin VB.Label Total 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "40"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2955
      TabIndex        =   13
      Top             =   5880
      Width           =   195
   End
   Begin VB.Image Mas 
      Height          =   135
      Index           =   3
      Left            =   3000
      Top             =   5250
      Width           =   135
   End
   Begin VB.Image Mas 
      Height          =   135
      Index           =   4
      Left            =   3000
      Top             =   5640
      Width           =   135
   End
   Begin VB.Image Mas 
      Height          =   135
      Index           =   2
      Left            =   3000
      Top             =   4830
      Width           =   135
   End
   Begin VB.Image Mas 
      Height          =   135
      Index           =   1
      Left            =   3000
      Top             =   4440
      Width           =   135
   End
   Begin VB.Image Mas 
      Height          =   135
      Index           =   0
      Left            =   3000
      Top             =   4050
      Width           =   135
   End
   Begin VB.Image Menos 
      Height          =   165
      Index           =   3
      Left            =   2445
      Top             =   5220
      Width           =   255
   End
   Begin VB.Image Menos 
      Height          =   135
      Index           =   4
      Left            =   2445
      Top             =   5640
      Width           =   255
   End
   Begin VB.Image Menos 
      Height          =   135
      Index           =   2
      Left            =   2445
      Top             =   4800
      Width           =   255
   End
   Begin VB.Image Menos 
      Height          =   135
      Index           =   1
      Left            =   2445
      Top             =   4440
      Width           =   255
   End
   Begin VB.Image Menos 
      Height          =   135
      Index           =   0
      Left            =   2445
      Top             =   4080
      Width           =   255
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   10440
      Top             =   6030
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   9720
      Top             =   6030
      Width           =   255
   End
   Begin VB.Label Informacion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   975
      Left            =   4200
      TabIndex        =   10
      Top             =   6000
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   3120
      Left            =   4935
      Stretch         =   -1  'True
      Top             =   2670
      Width           =   2475
   End
   Begin VB.Label Puntos 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7200
      TabIndex        =   9
      Top             =   8535
      Width           =   270
   End
   Begin VB.Image boton 
      Height          =   375
      Index           =   1
      Left            =   240
      MouseIcon       =   "frmCrearPersonaje.frx":0029
      MousePointer    =   99  'Custom
      Top             =   7680
      Width           =   1365
   End
   Begin VB.Image boton 
      Height          =   330
      Index           =   0
      Left            =   10440
      MouseIcon       =   "frmCrearPersonaje.frx":017B
      MousePointer    =   99  'Custom
      Top             =   7680
      Width           =   1320
   End
   Begin VB.Label lbConstitucion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2715
      TabIndex        =   5
      Top             =   5595
      Width           =   225
   End
   Begin VB.Label lbInteligencia 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2715
      TabIndex        =   4
      Top             =   4815
      Width           =   210
   End
   Begin VB.Label lbCarisma 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2715
      TabIndex        =   3
      Top             =   5205
      Width           =   225
   End
   Begin VB.Label lbAgilidad 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2715
      TabIndex        =   2
      Top             =   4410
      Width           =   225
   End
   Begin VB.Label lbFuerza 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2715
      TabIndex        =   1
      Top             =   4020
      Width           =   210
   End
End
Attribute VB_Name = "frmCrearPersonaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public SkillPoints As Byte

Function CheckData() As Boolean

If UserRaza = 0 Then
    MsgBox "Seleccione la raza del personaje."
    Informacion.Caption = "Seleccione la raza del personaje."
    Exit Function
End If

If UserSexo = 0 Then
    MsgBox "Seleccione el sexo del personaje."
    Informacion.Caption = "Seleccione el sexo del personaje."
    Exit Function
End If

If UserClase = 0 Then
    MsgBox "Seleccione la clase del personaje."
    Informacion.Caption = "Seleccione la clase del personaje."
    Exit Function
End If

If Not Total.Caption = 0 Then
    MsgBox "Asigne los atributos del personaje."
    Informacion.Caption = "Asigne los atributos del personaje."
    Exit Function
End If

If txtNombre.Text = "" Then
    MsgBox "Su personaje debe de tener un nombre."
    Informacion.Caption = "Su personaje debe de tener un nombre."
    Exit Function
End If

Dim i As Integer
For i = 1 To NUMATRIBUTOS
    If UserAtributos(i) = 0 Then
        MsgBox "Los atributos del personaje son invalidos."
        Informacion.Caption = "Los atributos del personaje son invalidos."
        Exit Function
    End If
Next i

If Len(UserName) > 30 Then
    MsgBox ("El nombre debe tener menos de 30 letras.")
    Exit Function
End If

If Cabeza.Text = "" Then
    MsgBox "Seleccione la cabeza de su personaje."
    Exit Function
End If

CheckData = True


End Function

Private Sub boton_Click(Index As Integer)
    Call General_Set_Wav(SND_CLICK)
    
    Select Case Index
        Case 0
            
            UserName = txtNombre.Text
            
            If Right$(UserName, 1) = " " Then
                UserName = RTrim$(UserName)
                MsgBox "Nombre invalido, se han removido los espacios al final del nombre"
            End If
            
            UserRaza = lstRaza.ListIndex + 1
            UserSexo = lstGenero.ListIndex + 1
            UserClase = lstProfesion.ListIndex + 1
            
            UserAtributos(1) = lbFuerza.Caption
            UserAtributos(2) = lbAgilidad.Caption
            UserAtributos(3) = lbInteligencia.Caption
            UserAtributos(4) = lbCarisma.Caption
            UserAtributos(5) = lbConstitucion.Caption
            
            If Not CheckData Then Exit Sub
            
            EstadoLogin = E_MODO.CrearNuevoPj
            Informacion.Caption = "Espere unos segundo, se esta tramitando la información de su nuevo personaje."
                
            If frmMain.Winsock1.State <> sckConnected Then
                frmMensaje.msg.Caption = "Error: Se ha perdido la conexion con el server."
                Informacion.Caption = "Se ha perdido la conexion con el server."
            Else
                PJName = txtNombre.Text
                Call Login
            End If
            
        Case 1
            If Not Audio.MusicActivated = False Then
                General_Set_Song 2, True
            End If
            
            Unload Me
            frmCuenta.Visible = True
    End Select
End Sub


Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single

Randomize Timer

RandomNumber = (UpperBound - LowerBound + 1) * Rnd + LowerBound
If RandomNumber > UpperBound Then RandomNumber = UpperBound

End Function

Private Sub cabeza_Click()
    MiCabeza = Val(Cabeza.List(Cabeza.ListIndex))
    
    Call DrawGrhtoHdc(frmCrearPersonaje.PlayerView.hDC, HeadData(MiCabeza).Head(3).GrhIndex, 0, 0, False)
    frmCrearPersonaje.PlayerView.Refresh
End Sub

Private Sub Form_Load()
Me.Caption = Form_Caption
Me.Picture = General_Load_Picture_From_Resource("40.gif")

Dim i As Integer
lstProfesion.Clear
For i = LBound(ListaClases) To UBound(ListaClases)
    lstProfesion.AddItem ListaClases(i)
Next i

lstRaza.Clear

For i = LBound(ListaRazas()) To UBound(ListaRazas())
    lstRaza.AddItem ListaRazas(i)
Next i


lstProfesion.Clear

For i = LBound(ListaClases()) To UBound(ListaClases())
    lstProfesion.AddItem ListaClases(i)
Next i

lstProfesion.ListIndex = 1

Image1.Picture = General_Load_Picture_From_Resource(lstProfesion.ListIndex & ".gif")
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = vbLeftButton) Then Call Auto_Drag(Me.hwnd)
End Sub
Private Sub Image3_Click()
Informacion.Caption = "De la cabeza que elijas sera la cuales todos veran dentro del juego."
If Cabeza.ListIndex < Cabeza.ListCount - 1 Then Cabeza.ListIndex = Cabeza.ListIndex + 1
End Sub

Private Sub Image2_Click()
Informacion.Caption = "De la cabeza que elijas sera la cuales todos veran dentro del juego."
If Cabeza.ListIndex > 0 Then Cabeza.ListIndex = Cabeza.ListIndex - 1
End Sub

Private Sub lstProfesion_Click()
On Error Resume Next
    UserClase = lstProfesion.ListIndex + 1
    Informacion.Caption = "La clase influirá en las características principales que tenga tu personaje, asi como en las magias e items que podrá utilizar. Las estrellas que ves abajo te mostrarán en qué habilidades se destaca la misma."
    Image1.Picture = General_Load_Picture_From_Resource(lstProfesion.ListIndex + 1 & ".gif")
End Sub
Private Sub lstGenero_Click()
UserSexo = lstGenero.ListIndex + 1
Informacion.Caption = "Indica si el personaje será masculino o femenino. Esto influye en los items que podrá equipar."
Call DameOpciones
End Sub
Private Sub lstRaza_Click()
UserRaza = lstRaza.ListIndex + 1
Informacion.Caption = "De la raza que elijas dependerá cómo se modifiquen los atributos. Podés cambiar de raza para poder visualizar cómo se modifican los distintos atributos."
Call DameOpciones

Select Case (lstRaza.List(lstRaza.ListIndex))
    Case Is = "Humano"
        ModFuerza.Caption = "+1"
        ModConstitucion.Caption = "+2"
        ModAgilidad.Caption = "+1"
        ModInteligencia.Caption = "0"
        ModCarisma.Caption = "0"
    Case Is = "Elfo"
        ModFuerza.Caption = "-1"
        ModConstitucion.Caption = "+1"
        ModAgilidad.Caption = "+4"
        ModInteligencia.Caption = "+2"
        ModCarisma.Caption = "+2"
    Case Is = "Elfo Oscuro"
        ModFuerza.Caption = "+2"
        ModConstitucion.Caption = "0"
        ModAgilidad.Caption = "+2"
        ModInteligencia.Caption = "+2"
        ModCarisma.Caption = "-3"
    Case Is = "Enano"
        ModFuerza.Caption = "+3"
        ModConstitucion.Caption = "+3"
        ModAgilidad.Caption = "-1"
        ModInteligencia.Caption = "-6"
        ModCarisma.Caption = "-2"
    Case Is = "Gnomo"
        ModFuerza.Caption = "-4"
        ModAgilidad.Caption = "+3"
        ModInteligencia.Caption = "+3"
        ModCarisma.Caption = "+1"
        ModConstitucion.Caption = "0"
    Case Is = "Orco"
        ModFuerza.Caption = "+ 5"
        ModConstitucion.Caption = "+ 3"
        ModAgilidad.Caption = "- 2"
        ModInteligencia.Caption = "- 6"
        ModCarisma.Caption = "- 2"
End Select
End Sub

Private Sub txtNombre_GotFocus()
Informacion.Caption = "Sea cuidadoso al seleccionar el nombre de su personaje, Winter AO es un juego de rol, un mundo magico y fantastico, si selecciona un nombre obsceno o con connotación politica los administradores borrarán su personaje y no habrá ninguna posibilidad de recuperarlo."
End Sub

Private Sub Mas_Click(Index As Integer)
Select Case Index
    Case 0
        If lbFuerza.Caption < 18 And Total.Caption > 0 Then
        lbFuerza.Caption = lbFuerza.Caption + 1
        Total.Caption = Total.Caption - 1
        End If
    Case 1
        If lbAgilidad.Caption < 18 And Total.Caption > 0 Then
        lbAgilidad.Caption = lbAgilidad.Caption + 1
        Total.Caption = Total.Caption - 1
        End If
    Case 2
        If lbInteligencia.Caption < 18 And Total.Caption > 0 Then
        lbInteligencia.Caption = lbInteligencia.Caption + 1
        Total.Caption = Total.Caption - 1
        End If
    Case 3
        If lbCarisma.Caption < 18 And Total.Caption > 0 Then
        lbCarisma.Caption = lbCarisma.Caption + 1
        Total.Caption = Total.Caption - 1
        End If
    Case 4
        If lbConstitucion.Caption < 18 And Total.Caption > 0 Then
        lbConstitucion.Caption = lbConstitucion.Caption + 1
        Total.Caption = Total.Caption - 1
        End If
    End Select
End Sub
Private Sub Menos_Click(Index As Integer)
Select Case Index
    Case 0
        If Total.Caption = "40" Then Exit Sub
        If lbFuerza.Caption > 6 Then
        lbFuerza.Caption = lbFuerza.Caption - 1
        Total.Caption = Total.Caption + 1
        End If
    Case 1
        If Total.Caption = "40" Then Exit Sub
        If lbAgilidad.Caption > 6 Then
        lbAgilidad.Caption = lbAgilidad.Caption - 1
        Total.Caption = Total.Caption + 1
        End If
    Case 2
        If Total.Caption = "40" Then Exit Sub
        If lbInteligencia.Caption > 6 Then
        lbInteligencia.Caption = lbInteligencia.Caption - 1
        Total.Caption = Total.Caption + 1
        End If
    Case 3
        If Total.Caption = "40" Then Exit Sub
        If lbCarisma.Caption > 6 Then
        lbCarisma.Caption = lbCarisma.Caption - 1
        Total.Caption = Total.Caption + 1
        End If
    Case 4
        If Total.Caption = "40" Then Exit Sub
        If lbConstitucion.Caption > 6 Then
        lbConstitucion.Caption = lbConstitucion.Caption - 1
        Total.Caption = Total.Caption + 1
        End If
    End Select
End Sub
