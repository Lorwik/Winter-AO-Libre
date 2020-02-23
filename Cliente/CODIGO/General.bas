Attribute VB_Name = "Mod_General"
Option Explicit

Public Win2kXP As Boolean
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const GWL_EXSTYLE = -20
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2&
Private Const WS_EX_TRANSPARENT As Long = &H20&

Public Windows_Temp_Dir As String

Private OSInfo As OSVERSIONINFO
'************************
'To get OS version
Private Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128
End Type
Private Declare Function GetOSVersion Lib "kernel32" _
Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Const VER_PLATFORM_WIN32s As Long = 0&
Private Const VER_PLATFORM_WIN32_WINDOWS As Long = 1&
Private Const VER_PLATFORM_WIN32_NT As Long = 2&

Public iplst As String

Public bFogata As Boolean

'***********************************************************
'ScreenShoots
Public Const VK_SNAPSHOT = &H2C
Public Declare Sub keybd_event _
Lib "user32" ( _
ByVal bVk As Byte, _
ByVal bScan As Byte, _
ByVal dwFlags As Long, _
ByVal dwExtraInfo As Long)
'***********************************************************

Private lFrameTimer As Long

'***********************************************************
'PARTE DEL MODULO ESTADO MSN PROGRAMABLE
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
 
    Public Type COPYDATASTRUCT
      dwData As Long
      cbData As Long
      lpData As Long
    End Type
 
    Public Const WM_COPYDATA = &H4A
'/PARTE DEL MODULO ESTADO MSN PROGRAMABLE
'***********************************************************

Public Function DirGraficos() As String
    DirGraficos = App.Path & "\Graficos\"
End Function
Public Function DirSound() As String
    DirSound = App.Path & "\Wav\"
End Function
Public Function DirMP3() As String
    DirMP3 = App.Path & "\MP3\"
End Function
Public Function DirMapas() As String
    DirMapas = App.Path & "\Mapas\"
End Function

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    'Initialize randomizer
    Randomize Timer
    
    'Generate random number
    RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound
End Function

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, ByVal Text As String, Optional ByVal red As Integer = -1, Optional ByVal green As Integer, Optional ByVal blue As Integer, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False, Optional ByVal bCrLf As Boolean = False)
'******************************************
'Adds text to a Richtext box at the bottom.
'Automatically scrolls to new text.
'Text box MUST be multiline and have a 3D
'apperance!
'Pablo (ToxicWaste) 01/26/2007 : Now the list refeshes properly.
'Juan Martín Sotuyo Dodero (Maraxus) 03/29/2007 : Replaced ToxicWaste's code for extra performance.
'******************************************r
    With RichTextBox
        If Len(.Text) > 1000 Then
            'Get rid of first line
            .SelStart = InStr(1, .Text, vbCrLf) + 1
            .SelLength = Len(.Text) - .SelStart + 2
            .TextRTF = .SelRTF
        End If
        
        .SelStart = Len(RichTextBox.Text)
        .SelLength = 0
        .SelBold = bold
        .SelItalic = italic
        
        If Not red = -1 Then .SelColor = RGB(red, green, blue)
        
        .SelText = IIf(bCrLf, Text, Text & vbCrLf)
        
        'RichTextBox.Refresh
    End With
End Sub

'TODO : Never was sure this is really necessary....
'TODO : 08/03/2006 - (AlejoLp) Esto hay que volarlo...
Public Sub RefreshAllChars()
'*****************************************************************
'Goes through the charlist and replots all the characters on the map
'Used to make sure everyone is visible
'*****************************************************************
    Dim loopc As Long
    
    For loopc = 1 To LastChar
        If charlist(loopc).Active = 1 Then
            MapData(charlist(loopc).Pos.X, charlist(loopc).Pos.Y).CharIndex = loopc
        End If
    Next loopc
End Sub

Function AsciiValidos(ByVal cad As String) As Boolean
    Dim car As Byte
    Dim i As Long
    
    cad = LCase$(cad)
    
    For i = 1 To Len(cad)
        car = Asc(mid$(cad, i, 1))
        
        If ((car < 97 Or car > 122) Or car = Asc("º")) And (car <> 255) And (car <> 32) Then
            Exit Function
        End If
    Next i
    
    AsciiValidos = True
End Function

Function CheckUserData(ByVal checkemail As Boolean) As Boolean
    'Validamos los datos del user
    Dim loopc As Long
    Dim CharAscii As Integer
    
    If checkemail And UserEmail = "" Then
        MsgBox ("Dirección de email invalida")
        Exit Function
    End If
    
    If UserPassword = "" Then
        MsgBox ("Ingrese un password.")
        Exit Function
    End If
    
    For loopc = 1 To Len(UserPassword)
        CharAscii = Asc(mid$(UserPassword, loopc, 1))
        If Not LegalCharacter(CharAscii) Then
            MsgBox ("Password inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
            Exit Function
        End If
    Next loopc
    
    If UserName = "" Then
        MsgBox ("Ingrese un nombre de personaje.")
        Exit Function
    End If
    
    If Len(UserName) > 30 Then
        MsgBox ("El nombre debe tener menos de 30 letras.")
        Exit Function
    End If
    
    For loopc = 1 To Len(UserName)
        CharAscii = Asc(mid$(UserName, loopc, 1))
        If Not LegalCharacter(CharAscii) Then
            MsgBox ("Nombre inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
            Exit Function
        End If
    Next loopc
    
    CheckUserData = True
End Function

Sub UnloadAllForms()
On Error Resume Next
    Dim mifrm As Form
    
    For Each mifrm In Forms
        Unload mifrm
    Next
End Sub

Function LegalCharacter(ByVal KeyAscii As Integer) As Boolean
'*****************************************************************
'Only allow characters that are Win 95 filename compatible
'*****************************************************************
    'if backspace allow
    If KeyAscii = 8 Then
        LegalCharacter = True
        Exit Function
    End If
    
    'Only allow space, numbers, letters and special characters
    If KeyAscii < 32 Or KeyAscii = 44 Then
        Exit Function
    End If
    
    If KeyAscii > 126 Then
        Exit Function
    End If
    
    'Check for bad special characters in between
    If KeyAscii = 34 Or KeyAscii = 42 Or KeyAscii = 47 Or KeyAscii = 58 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii = 63 Or KeyAscii = 92 Or KeyAscii = 124 Then
        Exit Function
    End If
    
    'else everything is cool
    LegalCharacter = True
End Function

Sub SetConnected()
'*****************************************************************
'Sets the client to "Connect" mode
'*****************************************************************
    'Set Connected
    Connected = True
    
    'Unload the connect form
    Unload frmCuenta
    Unload frmCrearPersonaje
    Unload frmConnect
    
    frmMain.Label8.Caption = PJName
    Call SetMusicInfo("Jugando Winter AO Ultimate [" & PJName & "] [Nivel: " & UserLvl & "] [ www.aowinter.com.ar ]", "Games", "{1}{0}")
    'Load main form
    frmMain.Visible = True
    
    FPSFLAG = True
    
    IScombate = True
End Sub

Sub MoveTo(ByVal Direccion As E_Heading)
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modify Date: 06/28/2008
'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
' 06/03/2006: AlejoLp - Elimine las funciones Move[NSWE] y las converti a esta
' 12/08/2007: Tavo    - Si el usuario esta paralizado no se puede mover.
' 06/28/2008: NicoNZ - Saqué lo que impedía que si el usuario estaba paralizado se ejecute el sub.
'***************************************************
    Dim LegalOk As Boolean
    
    If Cartel Then Cartel = False
    
    If frmMain.SendTxt.Visible = True And Opciones.DeMove = True Then Exit Sub
    
    Select Case Direccion
        Case E_Heading.NORTH
            LegalOk = MoveToLegalPos(UserPos.X, UserPos.Y - 1)
        Case E_Heading.EAST
            LegalOk = MoveToLegalPos(UserPos.X + 1, UserPos.Y)
        Case E_Heading.SOUTH
            LegalOk = MoveToLegalPos(UserPos.X, UserPos.Y + 1)
        Case E_Heading.WEST
            LegalOk = MoveToLegalPos(UserPos.X - 1, UserPos.Y)
    End Select
    
    If LegalOk And Not UserParalizado Then
        If Not UserDescansar And Not UserMeditar Then
            Call WriteWalk(Direccion)
            MoveCharbyHead UserCharIndex, Direccion
            MoveScreen Direccion
            Call ActualizarMiniMapa(Direccion)
        Else
            If UserDescansar And Not UserAvisado Then
                UserAvisado = True
                Call WriteRest
            End If
            If UserMeditar And Not UserAvisado Then
                UserAvisado = True
                Call WriteMeditate
            End If
        End If
    Else
        If charlist(UserCharIndex).Heading <> Direccion Then
            Call WriteChangeHeading(Direccion)
        End If
    End If
    
    ' Update 3D sounds!
    Call Audio.MoveListener(UserPos.X, UserPos.Y)
End Sub

Sub RandomMove()
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modify Date: 06/03/2006
' 06/03/2006: AlejoLp - Ahora utiliza la funcion MoveTo
'***************************************************
    Call MoveTo(RandomNumber(NORTH, WEST))
End Sub

Public Sub CheckKeys()
'*****************************************************************
'Checks keys and respond
'*****************************************************************

    'No input allowed while Winter AO is not the active window
    If Not Multimod.IsAppActive() Then Exit Sub
    
    'No walking when in commerce or banking.
    If Comerciando Then Exit Sub
    
    'No walking while writting in the forum.
    If frmForo.Visible Then Exit Sub
    
    'If game is paused, abort movement.
    If pausa Then Exit Sub
    
    'Don't allow any these keys during movement..
    If UserMoving = 0 Then
        If Not UserEstupido Then
            'Move Up
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0 Then
                Call MoveTo(NORTH)
                frmMain.Coord.Caption = "[" & UserMap & ", " & UserPos.X & ", " & UserPos.Y & "]"
                Exit Sub
            End If
            
            'Move Right
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Then
                Call MoveTo(EAST)
                frmMain.Coord.Caption = "[" & UserMap & ", " & UserPos.X & ", " & UserPos.Y & "]"
                Exit Sub
            End If
        
            'Move down
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Then
                Call MoveTo(SOUTH)
                frmMain.Coord.Caption = "[" & UserMap & ", " & UserPos.X & ", " & UserPos.Y & "]"
                Exit Sub
            End If
        
            'Move left
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0 Then
                Call MoveTo(WEST)
                frmMain.Coord.Caption = "[" & UserMap & ", " & UserPos.X & ", " & UserPos.Y & "]"
                Exit Sub
            End If
            
            ' We haven't moved - Update 3D sounds!
            Call Audio.MoveListener(UserPos.X, UserPos.Y)
        Else
            Dim kp As Boolean
            kp = (GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0) Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0
            
            If kp Then
                Call RandomMove
            Else
                ' We haven't moved - Update 3D sounds!
                Call Audio.MoveListener(UserPos.X, UserPos.Y)
            End If
            frmMain.Coord.Caption = "(" & UserPos.X & "," & UserPos.Y & ")"
        End If
    End If
End Sub

'TODO : Si bien nunca estuvo allí, el mapa es algo independiente o a lo sumo dependiente del engine, no va acá!!!
Sub SwitchMap(ByVal Map As Integer, ByVal Dir_Map As String)
'**************************************************************
'Formato de mapas optimizado para reducir el espacio que ocupan.
'Diseñado y creado por Juan Martín Sotuyo Dodero (Maraxus) Y mejorado por Lorwik :P
'**************************************************************

    Dim Y As Long
    Dim X As Long
    Dim tempint As Integer
    Dim ByFlags As Byte
    Dim handle As Integer
    
    Dim TempLng As Byte
    Dim TempByte1 As Byte
    Dim TempByte2 As Byte
    Dim TempByte3 As Byte
    
    Particle_Group_Remove_All
    
    Dim i As Byte
    
    handle = FreeFile()
    
    Open Dir_Map For Binary As handle
    Seek handle, 1
            
    'map Header
    Get handle, , MapInfo.MapVersion
    Get handle, , MiCabecera
    Get handle, , tempint
    Get handle, , tempint
    Get handle, , tempint
    Get handle, , tempint
    
    'Load arrays
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
        
            For i = 0 To 3
                MapData(X, Y).light_value(i) = False
            Next i
    
            Get handle, , ByFlags
            
            MapData(X, Y).blocked = (ByFlags And 1)
            
            Get handle, , MapData(X, Y).Graphic(1).GrhIndex
            InitGrh MapData(X, Y).Graphic(1), MapData(X, Y).Graphic(1).GrhIndex
            
            'Layer 2 used?
            If ByFlags And 2 Then
                Get handle, , MapData(X, Y).Graphic(2).GrhIndex
                InitGrh MapData(X, Y).Graphic(2), MapData(X, Y).Graphic(2).GrhIndex
            Else
                MapData(X, Y).Graphic(2).GrhIndex = 0
            End If
                
            'Layer 3 used?
            If ByFlags And 4 Then
                Get handle, , MapData(X, Y).Graphic(3).GrhIndex
                InitGrh MapData(X, Y).Graphic(3), MapData(X, Y).Graphic(3).GrhIndex
            Else
                MapData(X, Y).Graphic(3).GrhIndex = 0
            End If
                
            'Layer 4 used?
            If ByFlags And 8 Then
                Get handle, , MapData(X, Y).Graphic(4).GrhIndex
                InitGrh MapData(X, Y).Graphic(4), MapData(X, Y).Graphic(4).GrhIndex
            Else
                MapData(X, Y).Graphic(4).GrhIndex = 0
            End If
            
            'Trigger used?
            If ByFlags And 16 Then
                Get handle, , MapData(X, Y).Trigger
            Else
                MapData(X, Y).Trigger = 0
            End If
            
            If ByFlags And 32 Then
               Get handle, , tempint
                MapData(X, Y).particle_group_index = General_Particle_Create(tempint, X, Y, -1)
            End If
            
            If ByFlags And 64 Then
                Get handle, , MapData(X, Y).base_light(0)
                Get handle, , MapData(X, Y).base_light(1)
                Get handle, , MapData(X, Y).base_light(2)
                Get handle, , MapData(X, Y).base_light(3)
                
                If MapData(X, Y).base_light(0) Then _
                    Get handle, , MapData(X, Y).light_value(0)
                
                If MapData(X, Y).base_light(1) Then _
                    Get handle, , MapData(X, Y).light_value(1)
                
                If MapData(X, Y).base_light(2) Then _
                    Get handle, , MapData(X, Y).light_value(2)
                
                If MapData(X, Y).base_light(3) Then _
                    Get handle, , MapData(X, Y).light_value(3)
            End If
            
            'Erase NPCs
            If MapData(X, Y).CharIndex > 0 Then
                Call EraseChar(MapData(X, Y).CharIndex)
            End If
            
            'Erase OBJs
            MapData(X, Y).ObjGrh.GrhIndex = 0
            
            MapData(X, Y).Blood.Active = 0
            MapData(X, Y).Blood.Grh.GrhIndex = 24470
            MapData(X, Y).Blood.LifeTime = 0
        Next X
    Next Y
    
    Close handle
    
    MapInfo.name = ""
    MapInfo.Music = ""
    
End Sub

Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Byte) As String
'*****************************************************************
'Gets a field from a delimited string
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/15/2004
'*****************************************************************
    Dim i As Long
    Dim LastPos As Long
    Dim CurrentPos As Long
    Dim delimiter As String * 1
    
    delimiter = Chr$(SepASCII)
    
    For i = 1 To Pos
        LastPos = CurrentPos
        CurrentPos = InStr(LastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i
    
    If CurrentPos = 0 Then
        ReadField = mid$(Text, LastPos + 1, Len(Text) - LastPos)
    Else
        ReadField = mid$(Text, LastPos + 1, CurrentPos - LastPos - 1)
    End If
End Function

Function FieldCount(ByRef Text As String, ByVal SepASCII As Byte) As Long
'*****************************************************************
'Gets the number of fields in a delimited string
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 07/29/2007
'*****************************************************************
'**************************************************************
'Author: Unknown
'Last Modify Date: Unknown
'**************************************************************
    Dim count As Long, curPos As Long, delimiter As String * 1
    
    If LenB(Text) = 0 Then Exit Function
    delimiter = Chr$(SepASCII)
    curPos = 0
    Do
        curPos = InStr(curPos + 1, Text, delimiter)
        count = count + 1
    Loop While curPos <> 0
    
    FieldCount = count
End Function

Function FileExist(ByVal file As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (Dir$(file, FileType) <> "")
End Function

Sub Main()

    Form_Caption = "WinterAO Ultimate " & App.Major & "." & App.Minor & "." & App.Revision
    
    '[Desactivar mientras se desarrolle]
    
    OriginalClientName = "Winter AO Ultimate"
    ClientName = App.EXEName
    DetectName = App.EXEName
    If ChangeName Then
        Call ClientOn
        End
    End If
   
    If GetVar(App.Path & "\init\Config.CFG", "UPDATE", "Y") = 1 Then
        Call WriteVar(App.Path & "\init\Config.CFG", "UPDATE", "Y", "0")
    Else
        MsgBox "¡Debes de ejecutar el cliente desde el Launcher!", vbInformation
        End
        Exit Sub
    End If
   
    If Debugger Then
        Call AntiDebugger
        End
    End If
   
    If FindPreviousInstance Then
        Call MsgBox("Winter AO ya esta corriendo! No es posible correr otra instancia del juego. Haga click en Aceptar para salir.", vbApplicationModal + vbInformation + vbOKOnly, "Error al ejecutar")
        End
    End If
   
    Call ModSeguridad.AntiShInitialize
   
    '[/Desactivar mientras se desarrolle]
    
    Win2kXP = General_Windows_Is_2000XP
    ChDrive App.Path
    ChDir App.Path
    General_Associate_Icon
    
    'Set Temporal Dir
    Windows_Temp_Dir = General_Get_Temp_Dir
    
    '*********************************************
    'Lorwik> Modificar para hacerlo seleccionable.
    NoRes = GetVar(App.Path & "\INIT\Config.cfg", "Video", "Res")
    Call Multimod.SetResolution
    '*********************************************
    
    'Establecemos el 0% de la carga
    Call frmCargando.establecerProgreso(0)
    
    Set Light = New clsLight
    DirectXInit
    'Ruta del modulo de carga
    Set SurfaceDB = New clsSurfaceManDyn
    
    frmCargando.Show
    frmCargando.Refresh
    
    'Lorwik> Mostamos la version y licencia en el frmconnect
    frmConnect.version = "Versión " & App.Major & "." & App.Minor & "." & App.Revision & " GNU/GPL"
    
    '******************Paquetes*******************************
    frmCargando.Estado.Caption = "Buscando Paquetes... "
    If General_File_Exists(App.Path & "\RECURSOS\tmp.WAO", vbNormal) Then
        Call MsgBox("Hay Actualizaciones para los recursos. El cliente se cerrará y se abrirá el launcher para aplicar la actualización.", vbOKOnly, "Cliente Desactualizado")
        Call Shell(App.Path & "\WinterAO Ultimate Launcher.exe", vbNormalFocus)
        End
    End If
    '******************Constantes*******************************
    
    frmCargando.Estado.Caption = "Iniciando constantes... "
    
    Call Fonts_Initializate
    Call InicializarNombres
    Call frmCargando.progresoConDelay(15)
    ' Initialize FONTTYPES
    Call Protocol.InitFonts
    Call frmCargando.progresoConDelay(20)
    UserMap = 1
    
    Opciones.SangreAct = Val(GetVar(App.Path & "\Init\Config.cfg", "Video", "Blood"))
    Opciones.AutoComandos = Val(GetVar(App.Path & "\Init\Config.cfg", "Otros", "AutoCommand"))
    Opciones.DeMove = Val(GetVar(App.Path & "\Init\Config.cfg", "Otros", "DeMove"))
        '********************Motor Grafico**************************
    frmCargando.Estado.Caption = "Iniciando motor gráfico... "
    
    Dim PREC As Byte
    PREC = GetVar(App.Path & "\Init\Config.cfg", "Video", "Precarga")
    If Not InitTileEngine(frmMain.MainViewPic.hwnd, 149, 13, 32, 32, 13, 17, PREC, 8, 8, 0.018) Then
        Call CloseClient
    End If
    
    Call Inventario.Initialize(DirectD3D8, frmMain.PicInv, MAX_INVENTORY_SLOTS)
    '***********************************************************
    
    '**********************DirectSound**************************
    frmCargando.Estado.Caption = "Iniciando DirectSound... "
    
    'Inicializamos el sonido
    Call Audio.Initialize(DirectX, frmMain.hwnd, Windows_Temp_Dir, Windows_Temp_Dir)
    'Enable / Disable audio
    Audio.MusicActivated = GetVar(App.Path & "\Init\Config.cfg", "Sound", "MP3")
    Audio.SoundActivated = GetVar(App.Path & "\Init\Config.cfg", "Sound", "Wav")
    Audio.SoundEffectsActivated = GetVar(App.Path & "\Init\Config.cfg", "Sound", "FXSound")
    Opciones.AmbientAct = Audio.SoundEffectsActivated
    Audio.SoundVolume = Val(GetVar(App.Path & "\Init\Config.cfg", "Sound", "SoundVolume"))
    
    If Audio.MusicActivated = True Then General_Set_Song 1, True
    
    Call frmCargando.progresoConDelay(45)
    '***********************************************************
    
    '*****************Animaciones Extra************************
    frmCargando.Estado.Caption = "Creando animaciones extra... "
    Call frmCargando.progresoConDelay(85)
    Call CargarAnimArmas
    Call CargarAnimEscudos
    Call CargarColores
    Call frmCargando.progresoConDelay(100)
    '***********************************************************
    Opciones.SangreAct = True
    frmCargando.Estado.Caption = "¡Bienvenido a Winter AO Ultimate!"
    
    'Give the user enough time to read the welcome text
    Call Sleep(1750)
    
    Unload frmCargando

    frmConnect.Visible = True
    
    'Inicialización de variables globales
    prgRun = True
    pausa = False
    
    'Set the intervals of timers
    Call MainTimer.SetInterval(TimersIndex.Attack, INT_ATTACK)
    Call MainTimer.SetInterval(TimersIndex.Work, INT_WORK)
    Call MainTimer.SetInterval(TimersIndex.UseItemWithU, INT_USEITEMU)
    Call MainTimer.SetInterval(TimersIndex.UseItemWithDblClick, INT_USEITEMDCK)
    Call MainTimer.SetInterval(TimersIndex.SendRPU, INT_SENTRPU)
    Call MainTimer.SetInterval(TimersIndex.CastSpell, INT_CAST_SPELL)
    Call MainTimer.SetInterval(TimersIndex.Arrows, INT_ARROWS)
    Call MainTimer.SetInterval(TimersIndex.CastAttack, INT_CAST_ATTACK)
    
   'Init timers
    Call MainTimer.Start(TimersIndex.Attack)
    Call MainTimer.Start(TimersIndex.Work)
    Call MainTimer.Start(TimersIndex.UseItemWithU)
    Call MainTimer.Start(TimersIndex.UseItemWithDblClick)
    Call MainTimer.Start(TimersIndex.SendRPU)
    Call MainTimer.Start(TimersIndex.CastSpell)
    Call MainTimer.Start(TimersIndex.Arrows)
    Call MainTimer.Start(TimersIndex.CastAttack)
    
    lFrameTimer = GetTickCount
        
Do While prgRun
        If frmMain.WindowState <> 1 And frmMain.Visible Then
            Call ShowNextFrame(frmMain.Top, frmMain.Left, frmMain.MouseX, frmMain.MouseY)
            
            'Play ambient sounds
            Call RenderSounds
            Call CheckKeys
        End If
                    
        'FPS Counter - mostramos las FPS
        If GetTickCount - lFrameTimer >= 1000 Then
            If FPSFLAG Then frmMain.lblFPS.Caption = Mod_TileEngine.FPS
        
            lFrameTimer = GetTickCount
        End If
               
        ' If there is anything to be sent, we send it
        Call FlushBuffer
        
        DoEvents
    Loop

    
    Call CloseClient
End Sub

Sub WriteVar(ByVal file As String, ByVal Main As String, ByVal var As String, ByVal value As String)
'*****************************************************************
'Writes a var to a text file
'*****************************************************************
    writeprivateprofilestring Main, var, value, file
End Sub

Function GetVar(ByVal file As String, ByVal Main As String, ByVal var As String) As String
'*****************************************************************
'Gets a Var from a text file
'*****************************************************************
    Dim sSpaces As String ' This will hold the input that the program will retrieve
    
    sSpaces = Space$(100) ' This tells the computer how long the longest string can be. If you want, you can change the number 100 to any number you wish
    
    GetPrivateProfileString Main, var, vbNullString, sSpaces, Len(sSpaces), file
    
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function
Public Function General_File_Exists(ByVal file_path As String, ByVal file_type As VbFileAttribute) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Checks to see if a file exists
'*****************************************************************
    If Dir(file_path, file_type) = "" Then
        General_File_Exists = False
    Else
        General_File_Exists = True
    End If
End Function
'[CODE 002]:MatuX
'
'  Función para chequear el email
'
'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba y evitar un chequeo innecesario
Public Function CheckMailString(ByVal sString As String) As Boolean
On Error GoTo errHnd
    Dim lPos  As Long
    Dim Lx    As Long
    Dim iAsc  As Integer
    
    '1er test: Busca un simbolo @
    lPos = InStr(sString, "@")
    If (lPos <> 0) Then
        '2do test: Busca un simbolo . después de @ + 1
        If Not (InStr(lPos, sString, ".", vbBinaryCompare) > lPos + 1) Then _
            Exit Function
        
        '3er test: Recorre todos los caracteres y los valída
        For Lx = 0 To Len(sString) - 1
            If Not (Lx = (lPos - 1)) Then   'No chequeamos la '@'
                iAsc = Asc(mid$(sString, (Lx + 1), 1))
                If Not CMSValidateChar_(iAsc) Then _
                    Exit Function
            End If
        Next Lx
        
        'Finale
        CheckMailString = True
    End If
errHnd:
End Function

'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba
Private Function CMSValidateChar_(ByVal iAsc As Integer) As Boolean
    CMSValidateChar_ = (iAsc >= 48 And iAsc <= 57) Or _
                        (iAsc >= 65 And iAsc <= 90) Or _
                        (iAsc >= 97 And iAsc <= 122) Or _
                        (iAsc = 95) Or (iAsc = 45) Or (iAsc = 46)
End Function

'TODO : como todo lo relativo a mapas, no tiene nada que hacer acá....
Function HayAgua(ByVal X As Integer, ByVal Y As Integer) As Boolean
    HayAgua = ((MapData(X, Y).Graphic(1).GrhIndex >= 1505 And MapData(X, Y).Graphic(1).GrhIndex <= 1520) Or _
            (MapData(X, Y).Graphic(1).GrhIndex >= 5665 And MapData(X, Y).Graphic(1).GrhIndex <= 5680) Or _
            (MapData(X, Y).Graphic(1).GrhIndex >= 13547 And MapData(X, Y).Graphic(1).GrhIndex <= 13562)) And _
                MapData(X, Y).Graphic(2).GrhIndex = 0
                
End Function

Public Sub ShowSendTxt()
    If Not frmCantidad.Visible Then
        frmMain.SendTxt.Visible = True
        frmMain.SendTxt.SetFocus
    End If
End Sub

''
' Removes all text from the console and dialogs

Public Sub CleanDialogs()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'Removes all text from the console and dialogs
'**************************************************************
    'Clean console and dialogs
    frmMain.RecTxt.Text = vbNullString
    
    Call DialogosClanes.RemoveDialogs
    
    Call Dialogos.RemoveAllDialogs
End Sub

Public Sub CloseClient()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 8/14/2007
'Frees all used resources, cleans up and leaves
'**************************************************************
    ' Allow new instances of the client to be opened
    Call Multimod.ReleaseInstance
    
    EngineRun = False
    frmCargando.Show
    frmCargando.Estado.Caption = "Liberando recursos..."
    
    Call Multimod.ResetResolution
    
    'Stop tile engine
    Call DeinitTileEngine
    
    'Destruimos los objetos públicos creados
    Set CustomKeys = Nothing
    Set SurfaceDB = Nothing
    Set Dialogos = Nothing
    Set Audio = Nothing
    Set Inventario = Nothing
    Set MainTimer = Nothing
    Set incomingData = Nothing
    Set outgoingData = Nothing
    
    'Establecemos el 100% de la carga
    Call frmCargando.establecerProgreso(100)
    
    Call UnloadAllForms

    'Establecemos el 0% de la carga
    Call frmCargando.progresoConDelay(0)
    End
End Sub

Public Function esGM(CharIndex As Integer) As Boolean
esGM = False
If charlist(CharIndex).priv >= 1 And charlist(CharIndex).priv <= 5 Or charlist(CharIndex).priv = 25 Then _
    esGM = True

End Function

Public Function getTagPosition(ByVal Nick As String) As Integer
Dim buf As Integer
buf = InStr(Nick, "<")
If buf > 0 Then
    getTagPosition = buf
    Exit Function
End If
buf = InStr(Nick, "[")
If buf > 0 Then
    getTagPosition = buf
    Exit Function
End If
getTagPosition = Len(Nick) + 2
End Function
Public Sub Relog()
   
    EstadoLogin = E_MODO.LoginCuenta
If frmMain.Winsock1.State <> sckClosed Then
            frmMain.Winsock1.Close
            DoEvents
        End If
       
        frmMain.Winsock1.Connect CurServerIp, CurServerPort
End Sub
'**************************************************************
'MiniMapa
Public Sub ActualizarMiniMapa(ByVal tHeading As E_Heading)
'Esta es la forma mas optima que se me ha ocurrido. Solo dibuja  vez.
    frmMain.UserM.Left = UserPos.X - 1
    frmMain.UserM.Top = UserPos.Y - 1
    frmMain.UserArea.Left = UserPos.X - 9
    frmMain.UserArea.Top = UserPos.Y - 8
End Sub
Public Sub DibujarMiniMapa()
Dim map_x, map_y, Capas As Byte
    For map_y = 1 To 100
        For map_x = 1 To 100
        For Capas = 1 To 2
            If MapData(map_x, map_y).Graphic(Capas).GrhIndex > 0 Then
                SetPixel frmMain.Minimap.hDC, map_x - 1, map_y - 1, GrhData(MapData(map_x, map_y).Graphic(Capas).GrhIndex).MiniMap_color
            End If
            If MapData(map_x, map_y).Graphic(4).GrhIndex > 0 Then
                SetPixel frmMain.Minimap.hDC, map_x - 1, map_y - 1, GrhData(MapData(map_x, map_y).Graphic(4).GrhIndex).MiniMap_color
            End If
        Next Capas
        Next map_x
    Next map_y
   
    frmMain.Minimap.Refresh
    Call ActualizarMiniMapa(0)
End Sub

'***********************************************************
Public Sub Make_Transparent_Richtext(ByVal hwnd As Long)

If Win2kXP Then _
    Call SetWindowLong(hwnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)

End Sub

Public Function General_Windows_Is_2000XP() As Boolean
'**************************************************************
'Author: Unknown
'Last Modify Date: Unknown
'Get the windows version
'**************************************************************
On Error GoTo ErrorHandler

Dim RetVal As Long

OSInfo.dwOSVersionInfoSize = Len(OSInfo)
RetVal = GetOSVersion(OSInfo)

If OSInfo.dwPlatformId = VER_PLATFORM_WIN32_NT And OSInfo.dwMajorVersion >= 5 Then
    General_Windows_Is_2000XP = True
Else
    General_Windows_Is_2000XP = False
End If

Exit Function

ErrorHandler:
    General_Windows_Is_2000XP = False

End Function

'***********************************************************
Public Function GenerateKey() As String
Dim i As Byte, tempstring As String
    For i = 1 To 6
        If RandomNumber(1, 2) = 1 Then
            tempstring = tempstring & RandomNumber(1, 9)
        Else
            tempstring = tempstring & IIf(RandomNumber(1, 2) = 1, LCase$(Chr(97 + Rnd() * 862150000 Mod 26)), UCase$(Chr(97 + Rnd() * 862150000 Mod 26)))
        End If
    Next i
            
    GenerateKey = tempstring
End Function

'*************************************************
'Renderizado de Personajes en el Crear:
Sub DameOpciones()
Dim i As Integer
If frmCrearPersonaje.lstGenero.ListIndex < 0 Or frmCrearPersonaje.lstRaza.ListIndex < 0 Then
    frmCrearPersonaje.Cabeza.Enabled = False
ElseIf frmCrearPersonaje.lstGenero.ListIndex <> -1 And frmCrearPersonaje.lstRaza.ListIndex <> -1 Then
    frmCrearPersonaje.Cabeza.Enabled = True
End If

frmCrearPersonaje.Cabeza.Clear
    
Select Case frmCrearPersonaje.lstGenero.List(frmCrearPersonaje.lstGenero.ListIndex)
   Case "Hombre"
        Select Case frmCrearPersonaje.lstRaza.List(frmCrearPersonaje.lstRaza.ListIndex)
            Case "Humano"
                For i = 1 To 30
                    frmCrearPersonaje.Cabeza.AddItem i
                Next i
            Case "Elfo"
                For i = 101 To 113
                    If i = 113 Then i = 201
                    frmCrearPersonaje.Cabeza.AddItem i
                Next i
            Case "Elfo Oscuro"
                For i = 202 To 209
                    frmCrearPersonaje.Cabeza.AddItem i
                Next i
            Case "Enano"
                For i = 301 To 305
                    frmCrearPersonaje.Cabeza.AddItem i
                Next i
            Case "Gnomo"
                For i = 401 To 406
                    frmCrearPersonaje.Cabeza.AddItem i
                Next i
            Case "Orco"
                For i = 516 To 525
                    frmCrearPersonaje.Cabeza.AddItem i
                Next i
            Case Else
                UserHead = 1
        End Select
   Case "Mujer"
        Select Case frmCrearPersonaje.lstRaza.List(frmCrearPersonaje.lstRaza.ListIndex)
            Case "Humano"
                For i = 70 To 76
                    frmCrearPersonaje.Cabeza.AddItem i
                Next i
            Case "Elfo"
                For i = 170 To 176
                    frmCrearPersonaje.Cabeza.AddItem i
                Next i
            Case "Elfo Oscuro"
                For i = 270 To 278
                    frmCrearPersonaje.Cabeza.AddItem i
                Next i
            Case "Gnomo"
                For i = 470 To 474
                    frmCrearPersonaje.Cabeza.AddItem i
                Next i
            Case "Enano"
                For i = 370 To 372
                    frmCrearPersonaje.Cabeza.AddItem i
                Next i

            Case "Orco"
                For i = 526 To 531
                    frmCrearPersonaje.Cabeza.AddItem i
                Next i
            Case Else
                frmCrearPersonaje.Cabeza.AddItem "70"
        End Select
End Select

frmCrearPersonaje.PlayerView.Cls

End Sub
Public Function ColorToDX8(ByVal Long_Color As Long) As Long
    Dim temp_color As String
    Dim red As Integer, blue As Integer, green As Integer
    
    temp_color = Hex(Long_Color)
    If Len(temp_color) < 6 Then
        'Give is 6 digits for easy RGB conversion.
        temp_color = String(6 - Len(temp_color), "0") + temp_color
    End If
    
    red = CLng("&H" + mid$(temp_color, 1, 2))
    green = CLng("&H" + mid$(temp_color, 3, 2))
    blue = CLng("&H" + mid$(temp_color, 5, 2))
    
    ColorToDX8 = D3DColorXRGB(red, green, blue)

End Function
Public Function General_Var_Get(ByVal file As String, ByVal Main As String, ByVal var As String) As String
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Get a var to from a text file
'*****************************************************************
    Dim l As Long
    Dim Char As String
    Dim sSpaces As String 'Input that the program will retrieve
    Dim szReturn As String 'Default value if the string is not found
   
    szReturn = ""
   
    sSpaces = Space$(5000)
   
    GetPrivateProfileString Main, var, szReturn, sSpaces, Len(sSpaces), file
   
    General_Var_Get = RTrim$(sSpaces)
    General_Var_Get = Left$(General_Var_Get, Len(General_Var_Get) - 1)
End Function
Public Function General_Field_Read(ByVal field_pos As Long, ByVal Text As String, ByVal delimiter As Byte) As String
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets a field from a delimited string
'*****************************************************************
    Dim i As Long
    Dim LastPos As Long
    Dim FieldNum As Long
   
    LastPos = 0
    FieldNum = 0
    For i = 1 To Len(Text)
        If delimiter = CByte(Asc(mid$(Text, i, 1))) Then
            FieldNum = FieldNum + 1
            If FieldNum = field_pos Then
                General_Field_Read = mid$(Text, LastPos + 1, (InStr(LastPos + 1, Text, Chr$(delimiter), vbTextCompare) - 1) - (LastPos))
                Exit Function
            End If
            LastPos = i
        End If
    Next i
    FieldNum = FieldNum + 1
    If FieldNum = field_pos Then
        General_Field_Read = mid$(Text, LastPos + 1)
    End If
End Function

'=-=-==-=-=-==-=-=-==-=-=-==-MODULO DEL ESTADO DE MSN=-=-==-=-=-==-=-=-==-=-=-==-=-=-==-=-=-==-=-=-==-=-=-==-
'=-=-==-=-=-==-=-=-==-=-=-==-=-=-==-=-=-==-=-=-==-=-=-==-=-=-==-=-=-==-=-=-==-=-=-==-=-=-==-=-=-==-=-=-==-=-=
Public Sub SetMusicInfo(ByRef r_sArtist As String, ByRef r_sAlbum As String, ByRef r_sTitle As String, Optional ByRef r_sWMContentID As String = vbNullString, Optional ByRef r_sFormat As String = "{0} - {1}", Optional ByRef r_bShow As Boolean = True)
 
       Dim udtData As COPYDATASTRUCT
       Dim sBuffer As String
       Dim hMSGRUI As Long
       
       'Total length can Not be longer Then 256 characters!
       'Any longer will simply be ignored by Messenger.
       sBuffer = "\0Games\0" & Abs(r_bShow) & "\0" & r_sFormat & "\0" & r_sArtist & "\0" & r_sTitle & "\0" & r_sAlbum & "\0" & r_sWMContentID & "\0" & vbNullChar
       
       udtData.dwData = &H547
       udtData.lpData = StrPtr(sBuffer)
       udtData.cbData = LenB(sBuffer)
       
       Do
           hMSGRUI = FindWindowEx(0&, hMSGRUI, "MsnMsgrUIManager", vbNullString)
           
           If (hMSGRUI > 0) Then
               Call SendMessage(hMSGRUI, WM_COPYDATA, 0, VarPtr(udtData))
           End If
           
       Loop Until (hMSGRUI = 0)
 
    End Sub
'***********************************************************
'=-=-==-=-=-==-=-=-==-=-=-==-/MODULO DEL ESTADO DE MSN=-=-==-=-=-==-=-=-==-=-=-==-=-=-==-=-=-==-=-=-==-=-=-==-
'=-=-==-=-=-==-=-=-==-=-=-==-=-=-==-=-=-==-=-=-==-=-=-==-=-=-==-=-=-==-=-=-==-=-=-==-=-=-==-=-=-==-=-=-==-=-=-

Public Function General_Set_Song(ByVal nMP3 As Byte, Modo As Boolean)

    If Audio.MusicActivated = True Then
        
        If Modo = True Then
                If Audio.GetActualMP3 <> nMP3 Then
                    
                    If Audio.GetActualMP3 <> 0 Then
                        Audio.MusicMP3Stop
                        Audio.MusicMP3Empty 'Lorwik> Liberamos el archivo para poderlo eliminar o nos tirará "Permiso Denegado"
                        Delete_File Windows_Temp_Dir & Audio.GetActualMP3 & ".mp3"
                    End If
                    
                    Audio.SetActualMP3 = nMP3
                    Audio.mp3file = Get_Extract(MP3, Audio.GetActualMP3 & ".mp3")
                    
                    'Primero el Play...
                    Audio.MusicMP3Play (Get_Extract(MP3, Audio.GetActualMP3 & ".mp3"))
                    'Y despues ajustamos el volumen :)
                    Audio.MusicMP3VolumeSet Val(GetVar(App.Path & "\Init\Config.cfg", "Sound", "MusicVolume"))
                End If
        Else
            Audio.MusicMP3Stop
            Audio.MusicMP3Empty 'Lorwik> Liberamos el archivo para poderlo eliminar o nos tirará "Permiso Denegado"
            If Audio.GetActualMP3 <> 0 Then _
                    Delete_File Windows_Temp_Dir & Audio.GetActualMP3 & ".mp3"
        End If
        
    End If
End Function

Public Function General_Set_Wav(ByVal TSnd As String, Optional ByVal X As Byte, Optional ByVal Y As Byte, Optional ByVal LoopSound As LoopStyle = Default)
Dim file As String
    'Play Sound
    file = Get_Extract(Wav, TSnd)
            
    Audio.PlayWave TSnd, X, Y, LoopSound
                
    Delete_File file
        
End Function
Public Sub Make_Transparent_Form(ByVal hwnd As Long, Optional ByVal bytOpacity As Byte = 128)

If Win2kXP Then
    Call SetWindowLong(hwnd, GWL_EXSTYLE, GetWindowLong(hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
    Call SetLayeredWindowAttributes(hwnd, 0, bytOpacity, LWA_ALPHA)
End If

End Sub

Public Sub UnMake_Transparent_Form(ByVal hwnd As Long)

If Win2kXP Then _
    Call SetWindowLong(hwnd, GWL_EXSTYLE, GetWindowLong(hwnd, GWL_EXSTYLE) And (Not WS_EX_TRANSPARENT))

End Sub

