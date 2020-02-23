VERSION 5.00
Begin VB.Form frmOpciones 
   BorderStyle     =   0  'None
   ClientHeight    =   6885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4875
   ClipControls    =   0   'False
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
   ForeColor       =   &H00000000&
   Icon            =   "frmOpciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmOpciones.frx":0152
   ScaleHeight     =   6885
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox ChkMove 
      Caption         =   "Auto Completar Comandos"
      Height          =   200
      Left            =   345
      TabIndex        =   12
      Top             =   3250
      Width           =   180
   End
   Begin VB.HScrollBar Slider1 
      Height          =   250
      Index           =   1
      LargeChange     =   15
      Left            =   1440
      Max             =   100
      SmallChange     =   2
      TabIndex        =   11
      Top             =   720
      Width           =   3015
   End
   Begin VB.HScrollBar Slider1 
      Height          =   250
      Index           =   0
      LargeChange     =   15
      Left            =   1440
      Max             =   0
      Min             =   -4000
      SmallChange     =   2
      TabIndex        =   10
      Top             =   420
      Width           =   3015
   End
   Begin VB.CheckBox Check3 
      Height          =   200
      Left            =   345
      TabIndex        =   9
      Top             =   2430
      Width           =   180
   End
   Begin VB.CheckBox Check2 
      Height          =   200
      Left            =   345
      TabIndex        =   5
      Top             =   1750
      Width           =   180
   End
   Begin VB.CheckBox Check1 
      ForeColor       =   &H00000000&
      Height          =   200
      Index           =   1
      Left            =   345
      TabIndex        =   4
      Top             =   750
      Width           =   180
   End
   Begin VB.CheckBox Check1 
      ForeColor       =   &H00000000&
      Height          =   200
      Index           =   0
      Left            =   345
      TabIndex        =   3
      Top             =   400
      Width           =   180
   End
   Begin VB.CheckBox ChkComandos 
      Caption         =   "Auto Completar Comandos"
      Height          =   200
      Left            =   345
      TabIndex        =   2
      Top             =   2990
      Width           =   180
   End
   Begin VB.CheckBox Check1 
      ForeColor       =   &H00000000&
      Height          =   200
      Index           =   2
      Left            =   345
      TabIndex        =   0
      Top             =   1100
      Width           =   180
   End
   Begin VB.Label PreMenos 
      BackStyle       =   0  'Transparent
      Caption         =   "<"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   8
      Top             =   2040
      Width           =   135
   End
   Begin VB.Label Precarga 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2220
      TabIndex        =   7
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label PreMas 
      BackStyle       =   0  'Transparent
      Caption         =   ">"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   2040
      Width           =   135
   End
   Begin VB.Image Command2 
      Height          =   255
      Left            =   3240
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Image cmdCustomKeys 
      Height          =   255
      Left            =   480
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label Informacion 
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   360
      TabIndex        =   1
      Top             =   4220
      Width           =   4215
   End
End
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private loading As Boolean

Private Sub Check1_Click(Index As Integer)
    If Not loading Then _
        Call General_Set_Wav(SND_CLICK)
    
    Select Case Index
        Case 0
            If Check1(0).value = vbUnchecked Then
                Call WriteVar(App.Path & "\Init\Config.cfg", "Sound", "MP3", 0)
                Audio.MusicActivated = False
                Slider1(0).Enabled = False
                Informacion.Caption = "Desactiva la música del juego"
            ElseIf Not Audio.MusicActivated Then  'Prevent the music from reloading
                Call WriteVar(App.Path & "\Init\Config.cfg", "Sound", "MP3", 1)
                Audio.MusicActivated = True
                Slider1(0).Enabled = True
                Slider1(0).value = GetVar(App.Path & "\Init\Config.cfg", "Sound", "MusicVolume")
                Informacion.Caption = "Activa la música del juego"
            End If
        
        Case 1
            If Check1(1).value = vbUnchecked Then
                Call WriteVar(App.Path & "\Init\Config.cfg", "Sound", "Wav", 0)
                Audio.SoundActivated = False
                frmMain.IsPlaying = PlayLoop.plNone
                Slider1(1).Enabled = False
                Informacion.Caption = "Desactiva los efectos especiales y sonidos del juego"
            Else
                Call WriteVar(App.Path & "\Init\Config.cfg", "Sound", "Wav", 1)
                Audio.SoundActivated = True
                Slider1(1).Enabled = True
                Slider1(1).value = GetVar(App.Path & "\Init\Config.cfg", "Sound", "SoundVolume")
                Call Ambient
                Informacion.Caption = "Activa los efectos especiales y sonidos del juego"
            End If
            
        Case 2
            If Check1(2).value = vbUnchecked Then
                Audio.SoundEffectsActivated = False
                Opciones.AmbientAct = False
                Call Audio.StopWave
                Call WriteVar(App.Path & "\Init\Config.cfg", "Sound", "FXSound", 0)
                Informacion.Caption = "Desactiva efectos especiales como el sonido de navegacion o la ambientacion del lugar."
            Else
                Audio.SoundEffectsActivated = True
                Opciones.AmbientAct = True
                Call Ambient
                Call WriteVar(App.Path & "\Init\Config.cfg", "Sound", "FXSound", 1)
                Informacion.Caption = "Activa efectos especiales como el sonido de navegacion o la ambientacion del lugar."
            End If
    End Select
End Sub

Private Sub Check2_Click()
    If Check2.value = vbUnchecked Then
        Opciones.VSync = False
        Call WriteVar(App.Path & "\Init\Config.cfg", "Video", "VSync", 0)
        MsgBox "Los cambios se realizaran cuando reinicie el cliente.", vbInformation
        Informacion.Caption = "La sincronización vertical permite ajustar los FPS a la frecuencia de tu monitor para asi mejorar la experiencia en el juego."
    Else
        Opciones.VSync = True
        Call WriteVar(App.Path & "\Init\Config.cfg", "Video", "VSync", 1)
    End If
End Sub

Private Sub Check3_Click()
    If Check3.value = vbUnchecked Then
        Opciones.SangreAct = False
        Call WriteVar(App.Path & "\Init\Config.cfg", "Video", "Blood", 0)
        MsgBox "Los cambios se realizaran cuando reinicie el cliente.", vbInformation
        Informacion.Caption = "Activa o Desactiva las manchas de sangre al ser golpeados."
    Else
        Opciones.SangreAct = True
        Call WriteVar(App.Path & "\Init\Config.cfg", "Video", "Blood", 1)
    End If
End Sub

Private Sub ChkComandos_Click()
If ChkComandos.value = vbUnchecked Then
    Opciones.AutoComandos = False
    Call WriteVar(App.Path & "\Init\Config.cfg", "Otros", "AutoCommand", 0)
Else
    Opciones.AutoComandos = True
    Call WriteVar(App.Path & "\Init\Config.cfg", "Otros", "AutoCommand", 1)
End If
Informacion.Caption = "Activa o Desactiva el autocompletar comandos."
End Sub

Private Sub ChkMove_Click()
If ChkMove.value = vbUnchecked Then
    Opciones.DeMove = False
    Call WriteVar(App.Path & "\Init\Config.cfg", "Otros", "DeMove", 0)
Else
    Opciones.DeMove = True
    Call WriteVar(App.Path & "\Init\Config.cfg", "Otros", "DeMove", 1)
End If
Informacion.Caption = "Desactiva el movimiento del personaje al escribir."
End Sub

Private Sub cmdCustomKeys_Click()
    If Not loading Then _
        Call General_Set_Wav(SND_CLICK)
    Call frmCustomKeys.Show(vbModal, Me)
    Informacion.Caption = "Configura a tu gusto las teclas del juego. Teclas clasicas por default."
End Sub

Private Sub Command2_Click()
    Unload Me
    frmMain.SetFocus
End Sub

Private Sub Form_Load()
    loading = True      'Prevent sounds when setting check's values
    
    Me.Picture = General_Load_Picture_From_Resource("13.gif")
    
    If Audio.MusicActivated Then
        Check1(0).value = vbChecked
        Slider1(0).Enabled = True
        Slider1(0).value = GetVar(App.Path & "\Init\Config.cfg", "Sound", "MusicVolume")
    Else
        Check1(0).value = vbUnchecked
        Slider1(0).Enabled = False
    End If
    
    If Audio.SoundActivated Then
        Check1(1).value = vbChecked
        Slider1(1).Enabled = True
        Slider1(1).value = GetVar(App.Path & "\Init\Config.cfg", "Sound", "SoundVolume")
    Else
        Check1(1).value = vbUnchecked
        Slider1(1).Enabled = False
    End If
    
    If Audio.SoundEffectsActivated Then
        Check1(2).value = vbChecked
    Else
        Check1(2).value = vbUnchecked
    End If
    
    If Opciones.VSync Then
        Check2.value = vbChecked
    Else
        Check2.value = vbUnchecked
    End If
    
    Precarga.Caption = GetVar(App.Path & "\Init\Config.cfg", "Video", "Precarga")
    
    If Opciones.SangreAct Then
        Check3.value = vbChecked
    Else
        Check3.value = vbUnchecked
    End If
    
    If Opciones.AutoComandos Then
        ChkComandos.value = vbChecked
    Else
        ChkComandos.value = vbUnchecked
    End If
    
    If Opciones.DeMove Then
        ChkMove.value = vbChecked
    Else
        ChkMove.value = vbUnchecked
    End If
    
    loading = False     'Enable sounds when setting check's values
End Sub


Private Sub PreMas_Click()
If Not Precarga.Caption = 5 Then
    Precarga.Caption = Precarga.Caption + 1
    Call WriteVar(App.Path & "\Init\Config.cfg", "Video", "Precarga", Precarga.Caption)
    Informacion.Caption = "Mayor nivel de precarga mejorara la apariencia del juego. (Si crees que tu pc no lo podria soportar, no aumentar al maximo)"
End If
End Sub

Private Sub PreMenos_Click()
If Not Precarga.Caption = 1 Then
    Precarga.Caption = Precarga.Caption - 1
    Call WriteVar(App.Path & "\Init\Config.cfg", "Video", "Precarga", Precarga.Caption)
    Informacion.Caption = "Menor nivel de precarga te ayudara a resolver problemas con el juego, ya que mostrara los graficos justos en el screen."
End If
End Sub

Private Sub Slider1_Change(Index As Integer)
    Select Case Index
        Case 0
            Call WriteVar(App.Path & "\Init\Config.cfg", "Sound", "MusicVolume", str(Slider1(0).value))
            Audio.MusicMP3VolumeSet Slider1(0).value
            Informacion.Caption = "Ajusta el volumen de la música del juego."
        Case 1
            Call WriteVar(App.Path & "\Init\Config.cfg", "Sound", "SoundVolume", Slider1(1).value)
            Audio.SoundVolume = GetVar(App.Path & "\Init\Config.cfg", "Sound", "SoundVolume")
            Informacion.Caption = "Ajusta el volumen de efectos especiales y sonidos del juego."
    End Select
End Sub

Private Sub Slider1_Scroll(Index As Integer)
    Select Case Index
        Case 0
            Call WriteVar(App.Path & "\Init\Config.cfg", "Sound", "MusicVolume", Slider1(0).value)
            Audio.MusicMP3VolumeSet Slider1(0).value
        Case 1
            Call WriteVar(App.Path & "\Init\Config.cfg", "Sound", "SoundVolume", Slider1(1).value)
            Audio.SoundVolume = GetVar(App.Path & "\Init\Config.cfg", "Sound", "SoundVolume")
    End Select
End Sub
