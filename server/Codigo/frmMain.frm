VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Winter-AO Ultimate Server"
   ClientHeight    =   4545
   ClientLeft      =   1950
   ClientTop       =   1815
   ClientWidth     =   5190
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0C0&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000004&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4545
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin VB.Timer EventoHora 
      Interval        =   60000
      Left            =   3480
      Top             =   1800
   End
   Begin VB.Timer tPiqueteC 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   3000
      Top             =   2280
   End
   Begin VB.Timer packetResend 
      Interval        =   10
      Left            =   3000
      Top             =   1800
   End
   Begin VB.CommandButton CMDDUMP 
      Caption         =   "Crear Log Crítico de Usuarios"
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   4200
      Width           =   2895
   End
   Begin VB.Timer Auditoria 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4440
      Top             =   1800
   End
   Begin VB.Timer GameTimer 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   3960
      Top             =   1800
   End
   Begin VB.Timer AutoSave 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   4440
      Top             =   2280
   End
   Begin VB.Timer npcataca 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   3480
      Top             =   2280
   End
   Begin VB.Timer TIMER_AI 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3960
      Top             =   2280
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Mensajea todos los clientes"
      Height          =   3375
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4935
      Begin VB.ListBox ListadoM 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2205
         ItemData        =   "frmMain.frx":1042
         Left            =   120
         List            =   "frmMain.frx":1044
         TabIndex        =   10
         Top             =   1080
         Width           =   4695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Enviar por consola"
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
         Left            =   2520
         TabIndex        =   5
         Top             =   720
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Enviar por Pop-Up"
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
         TabIndex        =   4
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox BroadMsg 
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Label lblLimpieza 
      BackStyle       =   0  'Transparent
      Caption         =   "Limpieza del mundo en:"
      Height          =   255
      Left            =   1800
      TabIndex        =   11
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de usuarios:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Clima 
      BackStyle       =   0  'Transparent
      Caption         =   "Clima:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Escuch 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Left            =   3720
      TabIndex        =   6
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label CantUsuarios 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   405
   End
   Begin VB.Label txStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   5520
      Width           =   45
   End
   Begin VB.Menu mnuControles 
      Caption         =   "Argentum"
      Begin VB.Menu mnuServidor 
         Caption         =   "Configuracion"
      End
      Begin VB.Menu mnuSystray 
         Caption         =   "Systray Servidor"
      End
      Begin VB.Menu mnuCerrar 
         Caption         =   "Cerrar Servidor"
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuMostrar 
         Caption         =   "&Mostrar"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ESCUCHADAS As Long

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
   
Const NIM_ADD = 0
Const NIM_DELETE = 2
Const NIF_MESSAGE = 1
Const NIF_ICON = 2
Const NIF_TIP = 4

Const WM_MOUSEMOVE = &H200
Const WM_LBUTTONDBLCLK = &H203
Const WM_RBUTTONUP = &H205

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function Shell_NotifyIconA Lib "SHELL32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Integer

Private Function setNOTIFYICONDATA(hWnd As Long, ID As Long, flags As Long, CallbackMessage As Long, Icon As Long, Tip As String) As NOTIFYICONDATA
    Dim nidTemp As NOTIFYICONDATA

    nidTemp.cbSize = Len(nidTemp)
    nidTemp.hWnd = hWnd
    nidTemp.uID = ID
    nidTemp.uFlags = flags
    nidTemp.uCallbackMessage = CallbackMessage
    nidTemp.hIcon = Icon
    nidTemp.szTip = Tip & Chr$(0)

    setNOTIFYICONDATA = nidTemp
End Function

Sub CheckIdleUser()
    Dim iUserIndex As Long
    
    For iUserIndex = 1 To MaxUsers
       'Conexion activa? y es un usuario loggeado?
       If UserList(iUserIndex).ConnID <> -1 And UserList(iUserIndex).flags.UserLogged Then
            'Actualiza el contador de inactividad
            UserList(iUserIndex).Counters.IdleCount = UserList(iUserIndex).Counters.IdleCount + 1
            If UserList(iUserIndex).Counters.IdleCount >= IdleLimit Then
                Call WriteShowMessageBox(iUserIndex, "Demasiado tiempo inactivo. Has sido desconectado..")
                'mato los comercios seguros
                If UserList(iUserIndex).ComUsu.DestUsu > 0 Then
                    If UserList(UserList(iUserIndex).ComUsu.DestUsu).flags.UserLogged Then
                        If UserList(UserList(iUserIndex).ComUsu.DestUsu).ComUsu.DestUsu = iUserIndex Then
                            Call WriteConsoleMsg(UserList(iUserIndex).ComUsu.DestUsu, "Comercio cancelado por el otro usuario.", FontTypeNames.FONTTYPE_TALK)
                            Call FinComerciarUsu(UserList(iUserIndex).ComUsu.DestUsu)
                            Call FlushBuffer(UserList(iUserIndex).ComUsu.DestUsu) 'flush the buffer to send the message right away
                        End If
                    End If
                    Call FinComerciarUsu(iUserIndex)
                End If
                Call Cerrar_Usuario(iUserIndex)
            End If
        End If
    Next iUserIndex
End Sub

Private Sub Auditoria_Timer()
On Error GoTo errhand
Static centinelSecs As Byte

centinelSecs = centinelSecs + 1

If centinelSecs = 5 Then
    'Every 5 seconds, we try to call the player's attention so it will report the code.
    Call modCentinela.CallUserAttention
    
    centinelSecs = 0
End If

Call PasarSegundo 'sistema de desconexion de 10 segs

Exit Sub

errhand:

Call LogError("Error en Timer Auditoria. Err: " & Err.description & " - " & Err.Number)
Resume Next

End Sub

Private Sub AutoSave_Timer()

On Error GoTo Errhandler
'fired every minute
Static Minutos As Long
Static MinutosLatsClean As Long
Static MinsPjesSave As Long

Dim i As Integer
Dim num As Long

Minutos = Minutos + 1

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
Call ModAreas.AreasOptimizacion
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

'Actualizamos el centinela
Call modCentinela.PasarMinutoCentinela

If Minutos = MinutosWs - 1 Then
    Call SendData(SendTarget.toall, 0, PrepareMessageConsoleMsg("Worldsave en 1 minuto ...", FontTypeNames.FONTTYPE_VENENO))
End If

If Minutos >= MinutosWs Then
    Call DoBackUp
    Call aClon.VaciarColeccion
    Minutos = 0
End If

If MinutosLatsClean = MinutosLatsClean - 1 Then
    Call SendData(SendTarget.toall, 0, PrepareMessageConsoleMsg("Limpieza de mundo en 5 minutos ...", FontTypeNames.FONTTYPE_TALK))
End If

lblLimpieza = "Limpieza del mundo en: " & LimpiezaTimerMinutos & " minutos."
 
If Not LimpiezaTimerMinutos = 0 Then
    LimpiezaTimerMinutos = LimpiezaTimerMinutos - 1
Else
    lblLimpieza = "Limpieza del mundo en: ¡Limpiando Mundo!"
    Call LimpiarMundo
End If
 
If LimpiezaTimerMinutos = 5 Then
Call SendData(SendTarget.toall, 0, PrepareMessageConsoleMsg("Servidor > Atencion, 5 minutos para limpieza del mundo. Tomar items del piso.", FontTypeNames.FONTTYPE_SERVER))
End If

Call PurgarPenas
Call Piedra
Call CheckIdleUser

'<<<<<-------- Log the number of users online ------>>>
Dim N As Integer
N = FreeFile()
Open App.Path & "\logs\numusers.log" For Output Shared As N
Print #N, NumUsers
Close #N
'<<<<<-------- Log the number of users online ------>>>

Exit Sub
Errhandler:
    Call LogError("Error en TimerAutoSave " & Err.Number & ": " & Err.description)
    Resume Next
End Sub


Private Sub EventoHora_Timer()
'*************Lorwik/Noche***************
'**********wwww.lwk-foros.net**********
 
'Sorteamos el clima
Call SortearClima
'*****************
    
'**********************Aprovechamos el timer para el sistema de AutoTorneo*****************************
Static AutoTorneo As Byte
AutoTorneo = AutoTorneo + 1
Select Case AutoTorneo
    Case 60
        '******Recibir Premio Castillo*********
        'Lorwik> Tambien aprovechamos este codigo para poner los premios del castillo xDD.
        Call DarPremioCastillos
        Call Recordatorios(RandomNumber(0, 4))
        '**************************************
    Case 84
        Call SendData(SendTarget.toall, 0, PrepareMessageConsoleMsg("AutoTorneo> 10 minutos para torneo automatico. ¡Preparanse!", FontTypeNames.FONTTYPE_GUILD))
    Case 89
        Call SendData(SendTarget.toall, 0, PrepareMessageConsoleMsg("AutoTorneo> 5 minutos para torneo automatico. ¡Preparense!", FontTypeNames.FONTTYPE_GUILD))
    Case 93
        Call SendData(SendTarget.toall, 0, PrepareMessageConsoleMsg("AutoTorneo> 1 minuto para torneo automatico. ¡Preparense!", FontTypeNames.FONTTYPE_GUILD))
    Case 94
        Call Auto_Torneos(RandomNumber(1, 5))
    Case 96
        If Torneo_Esperando = True Then
            Call CancelarAutoTorneo
            AutoTorneo = 2
        Else
            AutoTorneo = 2
        End If
End Select
'*******************************************************************************************************

'+++++++ Evento automático [MaxTus] +++++++
    If HappyHourAC = True And mid(Format(Time, "HH:MM:SS"), 4, 2) = 0 Then
        HappyHourAC = False
        Call SendData(SendTarget.toall, 0, PrepareMessageConsoleMsg("Eventos Automáticos> El evento de experiencia x2 ha finalizado", FontTypeNames.FONTTYPE_GMMSG))
    End If
                    
    Call HappyHourAzar
'+++++++                            +++++++

End Sub
Private Sub CMDDUMP_Click()
On Error Resume Next

Dim i As Integer
For i = 1 To MaxUsers
    Call LogCriticEvent(i & ") ConnID: " & UserList(i).ConnID & ". ConnidValida: " & UserList(i).ConnIDValida & " Name: " & UserList(i).Name & " UserLogged: " & UserList(i).flags.UserLogged)
Next i

Call LogCriticEvent("Lastuser: " & LastUser & " NextOpenUser: " & NextOpenUser)

End Sub

Private Sub Command1_Click()
Call SendData(SendTarget.toall, 0, PrepareMessageShowMessageBox(BroadMsg.Text))
Call PostMensaje("Servidor> " & BroadMsg.Text)
End Sub

Public Sub InitMain(ByVal f As Byte)

If f = 1 Then
    Call mnuSystray_Click
Else
    frmMain.Show
End If

End Sub

Private Sub Command2_Click()
Call SendData(SendTarget.toall, 0, PrepareMessageConsoleMsg("Servidor> " & BroadMsg.Text, FontTypeNames.FONTTYPE_SERVER))
Call PostMensaje("Servidor> " & BroadMsg.Text)
End Sub

Private Sub Eventos_Timer()

    
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
   
   If Not Visible Then
        Select Case X \ Screen.TwipsPerPixelX
                
            Case WM_LBUTTONDBLCLK
                WindowState = vbNormal
                Visible = True
                Dim hProcess As Long
                GetWindowThreadProcessId hWnd, hProcess
                AppActivate hProcess
            Case WM_RBUTTONUP
                hHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf AppHook, App.hInstance, App.ThreadID)
                PopupMenu mnuPopUp
                If hHook Then UnhookWindowsHookEx hHook: hHook = 0
        End Select
   End If
   
End Sub
Private Sub QuitarIconoSystray()
On Error Resume Next

'Borramos el icono del systray
Dim i As Integer
Dim nid As NOTIFYICONDATA

nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, vbNull, frmMain.Icon, "")

i = Shell_NotifyIconA(NIM_DELETE, nid)
    

End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

'Save stats!!!
Call Statistics.DumpStatistics

Call QuitarIconoSystray

#If UsarQueSocket = 1 Then
Call LimpiaWsApi
#ElseIf UsarQueSocket = 0 Then
Socket1.Cleanup
#ElseIf UsarQueSocket = 2 Then
Serv.Detener
#End If

Dim LoopC As Integer

For LoopC = 1 To MaxUsers
    If UserList(LoopC).ConnID <> -1 Then Call CloseSocket(LoopC)
Next

'Log
Dim N As Integer
N = FreeFile
Open App.Path & "\logs\Main.log" For Append Shared As #N
Print #N, Date & " " & Time & " server cerrado."
Close #N

End

Set SonidosMapas = Nothing

End Sub


Private Sub GameTimer_Timer()
    Dim iUserIndex As Long
    Dim bEnviarStats As Boolean
    Dim bEnviarAyS As Boolean
    
On Error GoTo hayerror
    
    '<<<<<< Procesa eventos de los usuarios >>>>>>
    For iUserIndex = 1 To MaxUsers 'LastUser
        With UserList(iUserIndex)
           'Conexion activa?
           If .ConnID <> -1 Then
                '¿User valido?
                
                If .ConnIDValida And .flags.UserLogged Then
                    
                    '[Alejo-18-5]
                    bEnviarStats = False
                    bEnviarAyS = False
                    
                    If .flags.Paralizado = 1 Then Call EfectoParalisisUser(iUserIndex)
                    If .flags.Ceguera = 1 Or .flags.Estupidez Then Call EfectoCegueEstu(iUserIndex)
                    
                     'MaxTus
                        If .flags.Resucitando <> 0 Then
                            .Counters.Resucitar = .Counters.Resucitar + 1
                            If .Counters.Resucitar >= IntervaloPuedeResucitar Then
                                .Counters.Resucitar = 0
                                .flags.Resucitando = 0
                                RevivirUsuario iUserIndex
                            End If
                        End If
                    
                    If .flags.Muerto = 0 Then
                        
                        '[Consejeros]
                        If (.flags.Privilegios And PlayerType.User) Then Call EfectoLava(iUserIndex)
                        If .flags.Desnudo <> 0 And (.flags.Privilegios And PlayerType.User) <> 0 Then Call EfectoFrio(iUserIndex)
                        If .flags.Meditando Then Call DoMeditar(iUserIndex)
                        If .flags.Envenenado <> 0 And (.flags.Privilegios And PlayerType.User) <> 0 Then Call EfectoVeneno(iUserIndex)
                        If .flags.AdminInvisible <> 1 Then
                            If .flags.invisible = 1 Then Call EfectoInvisibilidad(iUserIndex)
                            If .flags.Oculto = 1 Then Call DoPermanecerOculto(iUserIndex)
                        End If
                        If .flags.Mimetizado = 1 Then Call EfectoMimetismo(iUserIndex)
                        If .flags.Metamorfosis = 1 Then Call EfectoMetamorfosis(iUserIndex) 'Metamorfosis
                        'MaxTus
                        If .flags.Makro <> 0 Then
                            .Counters.Makro = .Counters.Makro + 1
                            If .Counters.Makro >= IntervaloPuedeMakrear Then
                                .Counters.Makro = 0
                                MakroTrabajo iUserIndex, .flags.Makro
                            End If
                        End If
                        
                        Call DuracionPociones(iUserIndex)
                        
                        Call HambreYSed(iUserIndex, bEnviarAyS)
                        
                        If .flags.Hambre = 0 And .flags.Sed = 0 Then
                                If Not .flags.Descansar Then
                                'No esta descansando
                                    
                                    Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloSinDescansar)
                                    If bEnviarStats Then
                                        Call WriteUpdateHP(iUserIndex)
                                        bEnviarStats = False
                                    End If
                                    Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloSinDescansar)
                                    If bEnviarStats Then
                                        Call WriteUpdateSta(iUserIndex)
                                        bEnviarStats = False
                                    End If
                                    
                                Else
                                'esta descansando
                                    
                                    Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloDescansar)
                                    If bEnviarStats Then
                                        Call WriteUpdateHP(iUserIndex)
                                        bEnviarStats = False
                                    End If
                                    Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloDescansar)
                                    If bEnviarStats Then
                                        Call WriteUpdateSta(iUserIndex)
                                        bEnviarStats = False
                                    End If
                                    'termina de descansar automaticamente
                                    If .Stats.MaxHP = .Stats.MinHP And .Stats.MaxSta = .Stats.MinSta Then
                                        Call WriteRestOK(iUserIndex)
                                        Call WriteConsoleMsg(iUserIndex, "Has terminado de descansar.", FontTypeNames.FONTTYPE_INFO)
                                        .flags.Descansar = False
                                    End If
                                    
                                End If
                        End If
                        
                        If bEnviarAyS Then Call WriteUpdateHungerAndThirst(iUserIndex)
                        
                        If .NroMascotas > 0 Then Call TiempoInvocacion(iUserIndex)
                    End If 'Muerto
                Else 'no esta logeado?
                    'Inactive players will be removed!
                    .Counters.IdleCount = .Counters.IdleCount + 1
                    If .Counters.IdleCount > IntervaloParaConexion Then
                        .Counters.IdleCount = 0
                        Call CloseSocket(iUserIndex)
                    End If
                End If 'UserLogged
                
                'If there is anything to be sent, we send it
                Call FlushBuffer(iUserIndex)
            End If
        End With
    Next iUserIndex
Exit Sub

hayerror:
    LogError ("Error en GameTimer: " & Err.description & " UserIndex = " & iUserIndex)
End Sub



Private Sub mnuCerrar_Click()


If MsgBox("¡¡Atencion!! Si cierra el servidor puede provocar la perdida de datos. ¿Desea hacerlo de todas maneras?", vbYesNo) = vbYes Then
    Dim f
    For Each f In Forms
        Unload f
    Next
End If

End Sub

Private Sub mnusalir_Click()
    Call mnuCerrar_Click
End Sub

Public Sub mnuMostrar_Click()
On Error Resume Next
    WindowState = vbNormal
    Form_MouseMove 0, 0, 7725, 0
End Sub


Private Sub mnuServidor_Click()
frmServidor.Visible = True
End Sub

Private Sub mnuSystray_Click()

Dim i As Integer
Dim S As String
Dim nid As NOTIFYICONDATA

S = "WINTER-AO SERVER"
nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, WM_MOUSEMOVE, frmMain.Icon, S)
i = Shell_NotifyIconA(NIM_ADD, nid)
    
If WindowState <> vbMinimized Then WindowState = vbMinimized
Visible = False

End Sub

Private Sub npcataca_Timer()

On Error Resume Next
Dim npc As Long

For npc = 1 To LastNPC
    Npclist(npc).CanAttack = 1
Next npc

End Sub

Private Sub packetResend_Timer()
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/01/07
'Attempts to resend to the user all data that may be enqueued.
'***************************************************
On Error GoTo Errhandler:
    Dim i As Long
    
    For i = 1 To MaxUsers
        If UserList(i).ConnIDValida Then
            If UserList(i).outgoingData.length > 0 Then
                Call EnviarDatosASlot(i, UserList(i).outgoingData.ReadASCIIStringFixed(UserList(i).outgoingData.length))
            End If
        End If
    Next i

Exit Sub

Errhandler:
    LogError ("Error en packetResend - Error: " & Err.Number & " - Desc: " & Err.description)
    Resume Next
End Sub


Private Sub TIMER_AI_Timer()

On Error GoTo ErrorHandler
Dim NpcIndex As Long
Dim X As Integer
Dim Y As Integer
Dim UseAI As Integer
Dim mapa As Integer
Dim e_p As Integer

'Barrin 29/9/03
If Not haciendoBK And Not EnPausa Then
    'Update NPCs
    For NpcIndex = 1 To LastNPC
        
        If Npclist(NpcIndex).flags.NPCActive Then 'Nos aseguramos que sea INTELIGENTE!
            If Npclist(NpcIndex).flags.Paralizado = 1 Then
                Call EfectoParalisisNpc(NpcIndex)
                Else
                    'Usamos AI si hay algun user en el mapa
                    If Npclist(NpcIndex).flags.Inmovilizado = 1 Then
                       Call EfectoParalisisNpc(NpcIndex)
                    End If
                    
                    mapa = Npclist(NpcIndex).Pos.map
                    
                    If mapa > 0 Then
                        If MapInfo(mapa).NumUsers > 0 Then
                            If Npclist(NpcIndex).Movement <> TipoAI.ESTATICO Then
                                Call NPCAI(NpcIndex)
                            End If
                        End If
                    End If
                End If
            End If
    Next NpcIndex
End If

Exit Sub

ErrorHandler:
    Call LogError("Error en TIMER_AI_Timer " & Npclist(NpcIndex).Name & " mapa:" & Npclist(NpcIndex).Pos.map)
    Call MuereNpc(NpcIndex, 0)
End Sub




Private Sub tPiqueteC_Timer()
    Dim NuevaA As Boolean
    Dim NuevoL As Boolean
    Dim GI As Integer
    
    Dim i As Long
    
On Error GoTo Errhandler
    For i = 1 To LastUser
        With UserList(i)
            If .flags.UserLogged Then
                If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = eTrigger.ANTIPIQUETE Then
                    .Counters.PiqueteC = .Counters.PiqueteC + 1
                    Call WriteConsoleMsg(i, "¡¡¡Estás obstruyendo la vía pública, muévete o serás encarcelado!!!", FontTypeNames.FONTTYPE_INFO)
                    
                    If .Counters.PiqueteC > 23 Then
                        .Counters.PiqueteC = 0
                        Call Encarcelar(i, TIEMPO_CARCEL_PIQUETE)
                    End If
                Else
                    .Counters.PiqueteC = 0
                End If
                
                'ustedes se preguntaran que hace esto aca?
                'bueno la respuesta es simple: el codigo de AO es una mierda y encontrar
                'todos los puntos en los cuales la alineacion puede cambiar es un dolor de
                'huevos, asi que lo controlo aca, cada 6 segundos, lo cual es razonable
        
                GI = .GuildIndex
                If GI > 0 Then
                    NuevaA = False
                    NuevoL = False
                    If Not modGuilds.m_ValidarPermanencia(i, True, NuevaA, NuevoL) Then
                        Call WriteConsoleMsg(i, "Has sido expulsado del clan. ¡El clan ha sumado un punto de antifacción!", FontTypeNames.FONTTYPE_GUILD)
                    End If
                    If NuevaA Then
                        Call SendData(SendTarget.ToGuildMembers, GI, PrepareMessageConsoleMsg("¡El clan ha pasado a tener alineación neutral!", FontTypeNames.FONTTYPE_GUILD))
                        Call LogClanes("El clan cambio de alineacion!")
                    End If
                    If NuevoL Then
                        Call SendData(SendTarget.ToGuildMembers, GI, PrepareMessageConsoleMsg("¡El clan tiene un nuevo líder!", FontTypeNames.FONTTYPE_GUILD))
                        Call LogClanes("El clan tiene nuevo lider!")
                    End If
                End If
                
                Call FlushBuffer(i)
            End If
        End With
    Next i
Exit Sub

Errhandler:
    Call LogError("Error en tPiqueteC_Timer " & Err.Number & ": " & Err.description)
End Sub





'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''USO DEL CONTROL TCPSERV'''''''''''''''''''''''''''
'''''''''''''Compilar con UsarQueSocket = 3''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


#If UsarQueSocket = 3 Then

Private Sub TCPServ_Eror(ByVal Numero As Long, ByVal Descripcion As String)
    Call LogError("TCPSERVER SOCKET ERROR: " & Numero & "/" & Descripcion)
End Sub

Private Sub TCPServ_NuevaConn(ByVal ID As Long)
On Error GoTo errorHandlerNC

    ESCUCHADAS = ESCUCHADAS + 1
    Escuch.Caption = ESCUCHADAS
    
    Dim i As Integer
    
    Dim NewIndex As Integer
    NewIndex = NextOpenUser
    
    If NewIndex <= MaxUsers Then
        'call logindex(NewIndex, "******> Accept. ConnId: " & ID)
        
        TCPServ.SetDato ID, NewIndex
        
        If aDos.MaxConexiones(TCPServ.GetIP(ID)) Then
            Call aDos.RestarConexion(TCPServ.GetIP(ID))
            Call ResetUserSlot(NewIndex)
            Exit Sub
        End If

        If NewIndex > LastUser Then LastUser = NewIndex

        UserList(NewIndex).ConnID = ID
        UserList(NewIndex).ip = TCPServ.GetIP(ID)
        UserList(NewIndex).ConnIDValida = True
        Set UserList(NewIndex).CommandsBuffer = New CColaArray
        
        For i = 1 To BanIps.Count
            If BanIps.Item(i) = TCPServ.GetIP(ID) Then
                Call ResetUserSlot(NewIndex)
                Exit Sub
            End If
        Next i

    Else
        Call CloseSocket(NewIndex, True)
        LogCriticEvent ("NEWINDEX > MAXUSERS. IMPOSIBLE ALOCATEAR SOCKETS")
    End If

Exit Sub

errorHandlerNC:
Call LogError("TCPServer::NuevaConexion " & Err.description)
End Sub

Private Sub TCPServ_Close(ByVal ID As Long, ByVal MiDato As Long)
    On Error GoTo eh
    '' No cierro yo el socket. El on_close lo cierra por mi.
    'call logindex(MiDato, "******> Remote Close. ConnId: " & ID & " Midato: " & MiDato)
    Call CloseSocket(MiDato, False)
Exit Sub
eh:
    Call LogError("Ocurrio un error en el evento TCPServ_Close. ID/miDato:" & ID & "/" & MiDato)
End Sub

Private Sub TCPServ_Read(ByVal ID As Long, Datos As Variant, ByVal Cantidad As Long, ByVal MiDato As Long)
On Error GoTo errorh

With UserList(MiDato)
    Datos = StrConv(StrConv(Datos, vbUnicode), vbFromUnicode)
    
    Call .incomingData.WriteASCIIStringFixed(Datos)
    
    If .ConnID <> -1 Then
        Call HandleIncomingData(MiDato)
    Else
        Exit Sub
    End If
End With

Exit Sub

errorh:
Call LogError("Error socket read: " & MiDato & " dato:" & RD & " userlogged: " & UserList(MiDato).flags.UserLogged & " connid:" & UserList(MiDato).ConnID & " ID Parametro" & ID & " error:" & Err.description)

End Sub

#End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''FIN  USO DEL CONTROL TCPSERV'''''''''''''''''''''''''
'''''''''''''Compilar con UsarQueSocket = 3''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub Recordatorios(Index As Integer)
    Select Case Index
        Case 0
            Call SendData(SendTarget.toall, 0, PrepareMessageConsoleMsg("Servidor> Les recordamos que no esta permitido: insultos, propagandas de otros servidores, mandar denuncias pidiendole cosas a los GM's o llamandolos por su Nick, piquetes, u otras cosas que no estan permitidas, de lo contrario, serán advertidos. Para evitar todo este tipo de inconvenientes, visitese el manual en nuestra web accediendo a www.aowinter.com.ar o contactando con el email/msn soporte@aowinter.com.ar Muchas gracias.", FontTypeNames.FONTTYPE_SERVER))
        Case 1
            Call SendData(SendTarget.toall, 0, PrepareMessageConsoleMsg("Servidor> Les recordamos antes de enviar /GM consulten el manual en www.aowinter.com.ar/manual/ Muchas gracias.", FontTypeNames.FONTTYPE_SERVER))
        Case 2
            Call SendData(SendTarget.toall, 0, PrepareMessageConsoleMsg("Servidor> Les recordamos que si tiene alguna sugerencia envie ""/GM -> Sugerencia"". Muchas gracias.", FontTypeNames.FONTTYPE_SERVER))
        Case 3
            Call SendData(SendTarget.toall, 0, PrepareMessageConsoleMsg("Servidor> Les recordamos que si encuentra algun bug envie ""/GM -> Bug"". Muchas gracias.", FontTypeNames.FONTTYPE_SERVER))
        Case 4
            Call SendData(SendTarget.toall, 0, PrepareMessageConsoleMsg("Servidor> El staff les agradeceria que recomendasen el servidor a sus amigos/conocidos, asi como su promocion en internet. Muchas gracias.", FontTypeNames.FONTTYPE_SERVER))
    End Select
End Sub
