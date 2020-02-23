VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PathHelper - WinterAO Ultimate - By Lorwik"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5475
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Caption         =   "Solucion 5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   0
      TabIndex        =   28
      Top             =   960
      Visible         =   0   'False
      Width           =   5415
      Begin VB.CommandButton Command8 
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4440
         TabIndex        =   31
         Top             =   1720
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3240
         TabIndex        =   30
         Text            =   "0"
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   29
         Top             =   960
         Width           =   5175
      End
      Begin VB.Label Label19 
         Caption         =   $"Form1.frx":030A
         Height          =   615
         Left            =   120
         TabIndex        =   34
         Top             =   1320
         Width           =   5175
      End
      Begin VB.Label Label18 
         Caption         =   "1) Copiamos el enlace del Parche de Aqui:"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   720
         Width           =   5055
      End
      Begin VB.Label Label17 
         Caption         =   "Si despues de las anteriores soluciones no se soluciono tu problema, lo unico que queda es descargar el parche de forma manual."
         Height          =   495
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   5175
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Solucion 4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   0
      TabIndex        =   25
      Top             =   960
      Visible         =   0   'False
      Width           =   5415
      Begin VB.Label Label16 
         Caption         =   "Si con esto no se te soluciono pasa a la Solucion 5"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   1440
         TabIndex        =   27
         Top             =   1800
         Width           =   3855
      End
      Begin VB.Label Label15 
         Caption         =   $"Form1.frx":03D1
         Height          =   975
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   5175
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Solucion 3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   0
      TabIndex        =   20
      Top             =   960
      Visible         =   0   'False
      Width           =   5415
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Si no se te soluciono pasa a la Solucion 4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   24
         Top             =   1800
         Width           =   2775
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "2) Existen varias webs para hacer test de velocidad, con buscar un poco en google encontraremos miles."
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   1560
         Width           =   5295
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   $"Form1.frx":04F0
         Height          =   855
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   5295
      End
      Begin VB.Label Label11 
         Caption         =   "Debemos de asegurarno que estamos correctamente conectado a Internet, hay muchas formas de saber, aqui vamos a mostrar 2 de ellas."
         Height          =   615
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   5175
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Solucion 2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   0
      TabIndex        =   15
      Top             =   960
      Visible         =   0   'False
      Width           =   5415
      Begin VB.Label Label7 
         Caption         =   $"Form1.frx":05CA
         Height          =   615
         Left            =   120
         TabIndex        =   19
         Top             =   200
         Width           =   5055
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "1) Vamos a Inicio->Panel de Control-> Firewall-> Desactivar-> Aceptar"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   840
         Width           =   5295
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   $"Form1.frx":068C
         Height          =   975
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   5295
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   5400
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Si no se te soluciono pasa a la Solucion 3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   16
         Top             =   1850
         Width           =   2415
      End
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   5520
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   70
      TabIndex        =   10
      Top             =   3000
      Width           =   5415
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Solucion 5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   6
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Solucion 4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   5
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Solucion 3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Solucion 2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   720
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Solucion 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   5415
      Begin VB.CommandButton Command7 
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3840
         TabIndex        =   8
         Text            =   "0"
         Top             =   350
         Width           =   375
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   5175
         TabIndex        =   14
         Top             =   360
         Width           =   105
      End
      Begin VB.Label Label5 
         Caption         =   "Si tu problema no se soluciono, pasa a la Solucion 2."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   13
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "2) Pulsa aceptar, y una vez pulsado se iniciara el launcher, deberas de pulsar el boton ""Jugar"" o ""AutoUpdate""."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   5175
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Parche actual:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4280
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "1) Ingresa el numero del parche que causo el error: "
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Solucion 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label20 
      Caption         =   "Creado por Lorwik - www.aowinter.com.ar - www.rincondelao.com.ar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   120
      TabIndex        =   35
      Top             =   3240
      Width           =   5415
   End
   Begin VB.Label Label1 
      Caption         =   "Si tuviste algun problema al parchear el cliente, o no se te parcheo bien, sigue estas instrucciones:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5250
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Frame1.Visible = True
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = False
End Sub

Private Sub Command2_Click()
Frame2.Visible = True
Frame1.Visible = False
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = False
End Sub

Private Sub Command3_Click()
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = True
Frame4.Visible = False
Frame5.Visible = False
End Sub

Private Sub Command4_Click()
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = True
Frame5.Visible = False
End Sub

Private Sub Command5_Click()
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = True
End Sub

Private Sub Command8_Click()
   If Text3.Text < 0 Then
        MsgBox "El número debe ser mayor a 0"
        Exit Sub
    End If
    
    Call WriteVar(App.Path & "\INIT\" & "Config.cfg", "init", "X", Text3.Text)
   
End Sub

Private Sub Form_Load()
If Not FileExist(App.Path & "\Init\Config.cfg", vbArchive) Then
     MsgBox "No se ha encontrado el archivo Config.cfg"
     End
End If
Frame1.Visible = True
Label6.Caption = GetVar(App.Path & "\init\Config.cfg", "Init", "X")
    Dim Cargar As String
    Text2.Text = Inet1.OpenURL("http://www.aowinter.com.ar/update/url.txt")
End Sub
Private Sub Command6_Click()
End
End Sub
Private Sub Command7_Click()
   If Text1.Text < 0 Then
        MsgBox "El número debe ser mayor a 0"
        Exit Sub
    End If
    
    If Text1.Text > Label6.Caption Then
        MsgBox "El número debe ser menor a " & Label6.Caption
        Exit Sub
    End If
    
    Call WriteVar(App.Path & "\INIT\" & "Config.cfg", "init", "X", Text1.Text)
   
    Call ShellExecute(Me.hWnd, "open", "WinterAO Ultimate Launcher.exe", "", "", 1)
    End
End Sub

