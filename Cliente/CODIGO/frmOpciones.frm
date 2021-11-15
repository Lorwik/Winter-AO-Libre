VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOpciones 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Opciones del Juego"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   9120
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOpciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmOpciones.frx":0152
   ScaleHeight     =   4710
   ScaleWidth      =   9120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      Caption         =   "Registrar librerias necesarias para jugar"
      Height          =   195
      Left            =   4560
      TabIndex        =   37
      Top             =   600
      Width           =   3135
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Mp3"
      Height          =   375
      Left            =   2040
      TabIndex        =   36
      Top             =   270
      Value           =   1  'Checked
      Width           =   615
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Actualizador de Posición (Auto ""L"")"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   15
      Top             =   3240
      Width           =   4215
      Begin VB.CommandButton BotonCambiarTiempo 
         Caption         =   "Cambiar Tiempo"
         Height          =   255
         Left            =   1080
         TabIndex        =   21
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox Tiempo 
         Height          =   375
         Left            =   2520
         TabIndex        =   20
         Text            =   "60"
         Top             =   480
         Width           =   495
      End
      Begin VB.OptionButton ActPosicion 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Activado"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton DesactPosicion 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Desactivado"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Segundos."
         Height          =   255
         Index           =   0
         Left            =   3120
         TabIndex        =   19
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Actualizar posición cada"
         Height          =   255
         Left            =   1920
         TabIndex        =   18
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Información"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   2400
      Width           =   4215
      Begin VB.CommandButton Command1 
         Caption         =   "Foro de Discusión"
         Height          =   375
         Left            =   2640
         TabIndex        =   14
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdManual 
         Caption         =   "Manual Oficial"
         Height          =   375
         Left            =   1320
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton CmWeb 
         Caption         =   "Web Oficial"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Miscelaneos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3360
      Left            =   4440
      TabIndex        =   9
      Top             =   135
      Width           =   4575
      Begin VB.CheckBox Check2 
         Caption         =   "Dar 2 clicks a un usuario se abre el panel de Acciones"
         Height          =   195
         Left            =   240
         TabIndex        =   35
         Top             =   3000
         Value           =   1  'Checked
         Width           =   4215
      End
      Begin VB.CheckBox ActivarNoche 
         Caption         =   "Activar / Desactivar Efecto Noche"
         Height          =   195
         Left            =   240
         TabIndex        =   34
         Top             =   2680
         Width           =   3975
      End
      Begin VB.ComboBox msnState 
         Height          =   315
         ItemData        =   "frmOpciones.frx":159C
         Left            =   360
         List            =   "frmOpciones.frx":15AC
         TabIndex        =   33
         Text            =   "Jugando Winter-AO Return + [Nick] + [Nivel] + [Web]"
         Top             =   2280
         Width           =   3975
      End
      Begin VB.ComboBox visNom 
         Height          =   315
         ItemData        =   "frmOpciones.frx":1641
         Left            =   2640
         List            =   "frmOpciones.frx":164B
         TabIndex        =   30
         Text            =   "Activar"
         Top             =   1680
         Width           =   1695
      End
      Begin VB.ComboBox miniMap 
         Height          =   315
         ItemData        =   "frmOpciones.frx":1664
         Left            =   2640
         List            =   "frmOpciones.frx":166E
         TabIndex        =   29
         Text            =   "Activar"
         Top             =   1250
         Width           =   1695
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Intercambiar Botones del Mouse."
         Height          =   255
         Left            =   1680
         TabIndex        =   22
         Top             =   960
         Width           =   2775
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Configuración de Controles"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label6 
         Caption         =   "Estado del Msn:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   32
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Nombres:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   1680
         TabIndex        =   31
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "MiniMapa:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   1680
         TabIndex        =   28
         Top             =   1320
         Width           =   855
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1530
         Index           =   0
         Left            =   120
         Picture         =   "frmOpciones.frx":1687
         Top             =   720
         Width           =   1530
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1530
         Index           =   1
         Left            =   120
         Picture         =   "frmOpciones.frx":4828
         Top             =   720
         Width           =   1530
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Audio"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4215
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sonidos"
         Height          =   195
         Index           =   1
         Left            =   960
         TabIndex        =   8
         Top             =   240
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Música"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Value           =   1  'Checked
         Width           =   855
      End
      Begin MSComctlLib.Slider Slider2 
         Height          =   435
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   767
         _Version        =   393216
         Max             =   100
         SelStart        =   100
         Value           =   100
      End
      Begin MSComctlLib.Slider RigthSlider 
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1200
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   450
         _Version        =   393216
         Min             =   1
         Max             =   100
         SelStart        =   100
         Value           =   100
      End
      Begin MSComctlLib.Slider LeftSlider 
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1680
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   450
         _Version        =   393216
         Min             =   1
         Max             =   100
         SelStart        =   100
         Value           =   100
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Left:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Rigth:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   960
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Diálogos de clan"
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
      Height          =   855
      Left            =   4440
      TabIndex        =   1
      Top             =   3480
      Width           =   4575
      Begin VB.TextBox txtCantMensajes 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2640
         MaxLength       =   1
         TabIndex        =   4
         Text            =   "5"
         Top             =   360
         Width           =   450
      End
      Begin VB.OptionButton optPantalla 
         BackColor       =   &H00E0E0E0&
         Caption         =   "En pantalla"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   1320
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.OptionButton optConsola 
         BackColor       =   &H00E0E0E0&
         Caption         =   "En consola"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   60
         TabIndex        =   2
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "mensajes"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3240
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar"
      Height          =   225
      Left            =   1080
      MouseIcon       =   "frmOpciones.frx":79B6
      TabIndex        =   0
      Top             =   4440
      Width           =   6615
   End
End
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub ActivarNoche_Click()
If EfectosDiaY Then
EfectosDiaY = False
Call WriteVar(App.Path & "\init\config.ini", "Init", "Clima", 1)
Else
EfectosDiaY = True
Call WriteVar(App.Path & "\init\config.ini", "Init", "Clima", 0)
End If
End Sub

Private Sub ActPosicion_Click()
frmMain.ActualizadorPosicion.Enabled = True
End Sub

Private Sub BotonCambiarTiempo_Click()
frmMain.ActualizadorPosicion.Interval = Val(Tiempo.Text) * 1000
End Sub

Private Sub Command4_Click()
Shell "regsvr32 RICHTX32.OCX"
Shell "regsvr32 MSWINSCK.OCX"
Shell "regsvr32 COMCTL32.OCX"
Shell "regsvr32 MSCOMCTL.OCX"
Shell "regsvr32 COMDLG32.OCX"
Shell "regsvr32 CSWSK32.OCX"
Shell "regsvr32 DX7VB.DLL"
Shell "regsvr32 MSSTDFMT.DLL"
Shell "regsvr32 SCRRUN.DLL"
End Sub

Private Sub Check1_Click(Index As Integer)

Call Audio.PlayWave(SND_CLICK)

Select Case Index
    Case 0
        If Musica Then
            Musica = False
            Audio.StopMidi
        Else
            Musica = True
Call Extract_File2(Midi, App.Path & "\ARCHIVOS", CStr(currentMidi) & ".mid", Windows_Temp_Dir, False)
            Call Audio.PlayMIDI(CStr(currentMidi) & ".mid")
            Delete_File (Windows_Temp_Dir & CStr(currentMidi) & ".mid")
        End If
    Case 1
    
        If Sound Then
            Sound = False
            Call Audio.StopWave
            RainBufferIndex = 0
        Else
            Sound = True
        End If
End Select
End Sub
Private Sub Check2_Click()
If PanelsitoY Then
PanelsitoY = False
Call WriteVar(App.Path & "\init\config.ini", "Init", "Panelsito", 1)
Else
PanelsitoY = True
Call WriteVar(App.Path & "\init\config.ini", "Init", "Panelsito", 0)
End If
End Sub
Private Sub Check3_Click()
    If Check3 Then
        SwapMouseButton 1
        Image1(1).Visible = True
        Image1(0).Visible = False
    Else
        SwapMouseButton 0
        Image1(0).Visible = True
        Image1(1).Visible = False
    End If
End Sub
Private Sub CmWeb_Click()
Call ShellExecute(0, "Open", "http://winter-ao.com.ar/", "", App.Path, 0)
End Sub
Private Sub Command1_Click()
Call ShellExecute(0, "Open", "http://lwk-foros.com.ar/index.php?board=39.0", "", App.Path, 0)
End Sub
Private Sub Command2_Click()
Me.Visible = False
End Sub
Private Sub Command3_Click()
Call frmCustomKeys.Show(vbModeless, frmMain)
End Sub
Private Sub Check4_Click()
On Error Resume Next
If MPTres Then
            MPTres = False
            MP3P.stopMP3
            Call WriteVar(App.Path & "\init\config.ini", "Init", "MP3", "0")
        Else
            MPTres = True
            Call WriteVar(App.Path & "\init\config.ini", "Init", "MP3", "1")
            End If
End Sub
Private Sub DesactPosicion_Click()
frmMain.ActualizadorPosicion.Interval = 0
frmMain.ActualizadorPosicion.Enabled = False
End Sub
Private Sub Form_Load()
Call OpenMixer
Call ActualizaVolumen

    If Check3.value = True Then
        Image1(1).Visible = True
        Image1(0).Visible = False
    Else
        Image1(0).Visible = True
        Image1(1).Visible = False
    End If
    
    Get_Balance
    If GetVar(App.Path & "\init\config.ini", "Init", "Clima") = 1 Then
ActivarNoche.value = 1
EfectosDiaY = False
Else
ActivarNoche.value = 0
EfectosDiaY = True
End If

    If GetVar(App.Path & "\init\config.ini", "Init", "MP3") = 0 Then
Check4.value = 0
MPTres = False
Else
Check4.value = 1
MPTres = True
End If

    If GetVar(App.Path & "\init\config.ini", "Init", "panelsito") = 1 Then
Check2.value = 1
PanelsitoY = False
Else
Check2.value = 0
PanelsitoY = True
End If

End Sub
Private Sub miniMap_Click()
Select Case (miniMap.List(miniMap.listIndex))

    Case Is = "Desactivar"
     frmMain.miniMap.Visible = False
      frmMain.Label2.Visible = True
    Case Is = "Activar"
        frmMain.miniMap.Visible = True
        frmMain.Label2.Visible = False
End Select
End Sub
Private Sub msnState_Click()
Select Case (msnState.List(msnState.listIndex))

    Case Is = "Desactivar"
        Call SetMusicInfo("", "", "", "Games", , False)
            
    Case Is = "Jugando Winter-AO Return + [Nick]"
        Call SetMusicInfo("Jugando Winter-AO Return [" & UserName & "]", "Games", "{1}{0}")
            
    Case Is = "Jugando Winter-AO Return + [Nick] + [Nivel]"
        Call SetMusicInfo("Jugando Winter-AO Return [" & UserName & "] " & "[Nivel: " & UserLvl & "]", "Games", "{1}{0}")
            
    Case Is = "Jugando Winter-AO Return + [Nick] + [Nivel] + [Web]"
        Call SetMusicInfo("Jugando Winter-AO Return [" & UserName & "] [Nivel: " & UserLvl & "] [ www.winter-ao.com.ar ]", "Games", "{1}{0}")


End Select
End Sub

Private Sub optConsola_Click()
    DialogosClanes.Activo = False
End Sub

Private Sub optPantalla_Click()
    DialogosClanes.Activo = True
End Sub

Private Sub LeftSlider_Click()
Call Set_Balance
End Sub

Private Sub RigthSlider_Click()
Call Set_Balance
End Sub


Private Sub LeftSlider_Change()
Call Set_Balance
End Sub

Private Sub RigthSlider_Change()
Call Set_Balance
End Sub


Private Sub Slider2_Click()
Call ActualizaVolumen
End Sub

Private Sub txtCantMensajes_LostFocus()
    txtCantMensajes.Text = Trim$(txtCantMensajes.Text)
    If IsNumeric(txtCantMensajes.Text) Then
        DialogosClanes.CantidadDialogos = Trim$(txtCantMensajes.Text)
    Else
        txtCantMensajes.Text = 5
    End If
End Sub


Private Sub visNom_Click()
Select Case (visNom.List(visNom.listIndex))
    Case Is = "Desactivar"
        Nombres = False
    Case Is = "Activar"
        Nombres = True
End Select
End Sub
