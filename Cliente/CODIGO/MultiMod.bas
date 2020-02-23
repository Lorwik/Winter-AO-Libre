Attribute VB_Name = "Multimod"
'====================================================================================================
'MODULO CREADO PARA DIFERENTES SISTEMAS (asi no tengo que hacer 400 modulos para 3 codigos de mierda)
'NOTA: No usar "i" ni "j" para los For
'====================================================================================================
'*****************************
'MOD LIGHT
'*****************************
Public Type LightVertex
    type As Byte
    affected As Byte
End Type

Public Type DayLightType
    r As Byte
    g As Byte
    B As Byte
End Type

Public DayLightByte As DayLightType

Public TwinkLightByteHandle As Long

Public LightMap(1 To 100, 1 To 100) As LightVertex
'*****************************
'*****************************

'*****************************
'MOD RENDERIZADO DE FUENTES
'*****************************
Dim i As Integer
Dim j As Byte

Type Fuente
    Characters(32 To 255) As Long 'ASCII Characters
End Type

Private Fuentes() As Fuente
'*****************************
'*****************************

'*****************************
'MOD RESOLUCIÓN
'*****************************
Private Const CCDEVICENAME As Long = 32
Private Const CCFORMNAME As Long = 32
Private Const DM_BITSPERPEL As Long = &H40000
Private Const DM_PELSWIDTH As Long = &H80000
Private Const DM_PELSHEIGHT As Long = &H100000
Private Const DM_DISPLAYFREQUENCY As Long = &H400000
Private Const CDS_TEST As Long = &H4
Private Const ENUM_CURRENT_SETTINGS As Long = -1

Private Type typDevMODE
    dmDeviceName       As String * CCDEVICENAME
    dmSpecVersion      As Integer
    dmDriverVersion    As Integer
    dmSize             As Integer
    dmDriverExtra      As Integer
    dmFields           As Long
    dmOrientation      As Integer
    dmPaperSize        As Integer
    dmPaperLength      As Integer
    dmPaperWidth       As Integer
    dmScale            As Integer
    dmCopies           As Integer
    dmDefaultSource    As Integer
    dmPrintQuality     As Integer
    dmColor            As Integer
    dmDuplex           As Integer
    dmYResolution      As Integer
    dmTTOption         As Integer
    dmCollate          As Integer
    dmFormName         As String * CCFORMNAME
    dmUnusedPadding    As Integer
    dmBitsPerPel       As Integer
    dmPelsWidth        As Long
    dmPelsHeight       As Long
    dmDisplayFlags     As Long
    dmDisplayFrequency As Long
End Type

Private oldResHeight As Long
Private oldResWidth As Long
Private oldDepth As Integer
Private oldFrequency As Long
Private bNoResChange As Boolean


Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lptypDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lptypDevMode As Any, ByVal dwFlags As Long) As Long
'*****************************
'*****************************

'*****************************
'MOD CARTELES
'*****************************
Const XPosCartel = 125
Const YPosCartel = 100
Const MAXLONG = 40

'Carteles
Public Cartel As Boolean
Public Leyenda As String
Public LeyendaFormateada() As String
Public textura As Integer
'*****************************
'*****************************

'*****************************
'MOD PREVINSTANCE
'*****************************
'Declaration of the Win32 API function for creating /destroying a Mutex, and some types and constants.
Private Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (ByRef lpMutexAttributes As SECURITY_ATTRIBUTES, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
Private Declare Function ReleaseMutex Lib "kernel32" (ByVal hMutex As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Const ERROR_ALREADY_EXISTS = 183&

Private mutexHID As Long
'*****************************
'*****************************

'*****************************
'MOD COMANDOS
'*****************************
Public Const Commands As String = "Comerciar,Salir,GM,Boveda,Resucitar,Curar,Participar,Online,Meditar,Invocar,Descansar,Torneo,Creargrupo,Grupo,Fundarclan,Denunciar,Ping"
Public COM() As String

Public Mostrando As Boolean
'****************************
'****************************

'****************************
'MOD AREAS
'****************************
Public MinLimiteX As Integer
Public MaxLimiteX As Integer
Public MinLimiteY As Integer
Public MaxLimiteY As Integer
'***************************
'*****MOD Application*******
Private Declare Function GetActiveWindow Lib "user32" () As Long
'***************************
'***************************
'MOVER VENTANAS
'***************************
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    lParam As Any) As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
'***************************
'CLIMAS
'***************************
Type RGBClimax
    r As Byte
    g As Byte
    B As Byte
    A As Byte
End Type
 
Public ColorClimax As RGBClimax
'***************************

'To Get System32 Dir
Private Declare Function GetSystemDirectory _
    Lib "kernel32" _
    Alias "GetSystemDirectoryA" ( _
        ByVal lpBuffer As String, _
        ByVal nsize As Long) As Long

'To Copy Files
Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Declare Function CopyFileEx Lib "kernel32.dll" Alias "CopyFileExA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal lpProgressRoutine As Long, lpData As Any, ByRef pbCancel As Long, ByVal dwCopyFlags As Long) As Long

'*****************************
'MOVER VENTANAS
'*****************************
Public Sub Auto_Drag(ByVal hwnd As Long)
If NoRes = 0 Then Exit Sub
Call ReleaseCapture
Call SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&)
End Sub
'*****************************
'MOD LIGHT
'*****************************
Public Function LightValue(value As Integer) As Long

If value > 255 Then value = 255
value = value - TwinkLightByteHandle
LightValue = RGB(value, value, value)

End Function
Public Sub ClimaX()
'***************Lorwik/Noche*************************
'**********www.RincondelAO.com.ar********************
    Dim AmbientColor As D3DCOLORVALUE
    Dim Nw As D3DCOLORVALUE
    
    With AmbientColor
         'Si el usuario esta muerto mostramos otro color
        If UserEstado = 1 Or Zona = "DUNGEON" Then
            Call CalculateRGB(160, 160, 160, 1)
        Else
            'Mañana
            If Anocheceria = 0 Then
                Call CalculateRGB(230, 200, 200, 255)
            'MedioDia
            ElseIf Anocheceria = 1 Then
                Call CalculateRGB(255, 255, 255, 255)
            'Tarde
            ElseIf Anocheceria = 2 Then
                Call CalculateRGB(200, 200, 200, 255)
            'Noche
            ElseIf Anocheceria = 3 Then
                Call CalculateRGB(165, 165, 165, 1)
            End If

        End If
    End With
End Sub

Public Sub CalculateRGB(r As Byte, g As Byte, B As Byte, A As Byte)
With ColorClimax
    If .r < r Then
        .r = .r + 1
    Else
        .r = .r - 1
    End If
    If .B < B Then
        .B = .B + 1
    Else
        .B = .B - 1
    End If
    If .g < g Then
        .g = .g + 1
    Else
        .g = .g - 1
    End If
    If .A < A Then
        .A = .A + 1
    Else
        .A = .A - 1
    End If
    base_light = ARGB(.r, .g, .B, .A)
End With
End Sub
Public Sub SpeedCaballo()
    If UserEquitando Then
        engineBaseSpeed = 0.022
    Else
        engineBaseSpeed = 0.016
    End If
End Sub

Public Sub SetLight(X As Byte, Y As Byte)

With LightMap(X, Y)
    .affected = 4
    .type = 1
End With

AffectVertex X + 1, Y, 2
AffectVertex X - 1, Y, 2
AffectVertex X, Y - 1, 2
AffectVertex X, Y + 1, 2

AffectVertex X - 1, Y - 1, 1
AffectVertex X + 1, Y - 1, 1
AffectVertex X - 1, Y + 1, 1
AffectVertex X + 1, Y + 1, 1

End Sub

Public Sub AffectVertex(X As Byte, Y As Byte, value As Byte)

With LightMap(X, Y)
    If .affected < value Then .affected = value
    .type = 1
End With

End Sub
'*****************************
'*****************************

'*****************************
'MOD RENDERIZADO DE FUENTES
'*****************************
Public Sub Fonts_DeInit()
    Erase Fuentes()
    Exit Sub
End Sub

Public Sub Fonts_Initializate()
    Dim Leer As New clsIniReader
    Dim Num_Fuentes As Byte
    Dim file As String
    
    file = Get_Extract(Scripts, "Fuentes.ini")
    
    Leer.Initialize file
    
    Num_Fuentes = Val(Leer.GetValue("Fuentes", "Num_Fuentes"))
    ReDim Fuentes(1 To Num_Fuentes)
    
    
        For j = 1 To Num_Fuentes
            For i = 32 To 255
                Fuentes(j).Characters(i) = Val(Leer.GetValue("Fuentes", "Fuentes(" & j & ").Caracteres(" & i & ") "))
            Next i
        Next j
        
    Set Leer = Nothing
    
    Delete_File file

End Sub

Public Sub Fonts_Render_String(ByVal Text As String, ByVal X As Integer, ByVal Y As Integer, Optional ByVal c As Long = -1, Optional ByVal Font_Num As Byte = 1)
If Len(Text) = 0 Then Exit Sub
Dim color(0 To 3) As Long, GrhIndex As Long, Suma As Integer
    Long_To_RGB_List color(), c
    
    Suma = 0
    
    For i = 1 To Len(Text)
        GrhIndex = Fuentes(Font_Num).Characters(Asc(mid(Text, i, 1)))
        DDrawTransGrhIndextoSurface GrhIndex, X + Suma, Y, 0, color()
        Suma = Suma + GrhData(GrhIndex).pixelWidth - 2
    Next i
    
End Sub

Public Function Fonts_Render_String_Width(ByVal Text As String, Optional ByVal Font_Num As Byte = 1) As Integer
If Len(Text) = 0 Then
    Fonts_Render_String_Width = 0
    Exit Function
End If

Dim Suma As Integer
    Suma = 0
    
    For i = 1 To Len(Text)
        Suma = Suma + GrhData(Fuentes(Font_Num).Characters(Asc(mid(Text, i, 1)))).pixelWidth
    Next i
    
    Fonts_Render_String_Width = Suma
End Function
'*****************************
'*****************************

'*****************************
'MOD RESOLUCIÓN
'*****************************
Public Sub SetResolution()
'***************************************************
'Autor: Unknown
'Last Modification: 03/29/08
'Changes the display resolution if needed.
'Last Modified By: Juan Martín Sotuyo Dodero (Maraxus)
' 03/29/2008: Maraxus - Retrieves current settings storing display depth and frequency for proper restoration.
'***************************************************
    Dim lRes As Long
    Dim MidevM As typDevMODE
    Dim CambiarResolucion As Boolean
    
    lRes = EnumDisplaySettings(0, ENUM_CURRENT_SETTINGS, MidevM)
    
    oldResWidth = Screen.Width \ Screen.TwipsPerPixelX
    oldResHeight = Screen.Height \ Screen.TwipsPerPixelY
    
    If NoRes Then
        CambiarResolucion = (oldResWidth < 800 Or oldResHeight < 600)
    Else
        CambiarResolucion = (oldResWidth <> 800 Or oldResHeight <> 600)
    End If
    
    If CambiarResolucion Then
        
        With MidevM
            oldDepth = .dmBitsPerPel
            oldFrequency = .dmDisplayFrequency
            
            .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
            .dmPelsWidth = 800
            .dmPelsHeight = 600
            .dmBitsPerPel = 16
        End With
        
        lRes = ChangeDisplaySettings(MidevM, CDS_TEST)
    Else
        bNoResChange = True
    End If
End Sub

Public Sub ResetResolution()
'***************************************************
'Autor: Unknown
'Last Modification: 03/29/08
'Changes the display resolution if needed.
'Last Modified By: Juan Martín Sotuyo Dodero (Maraxus)
' 03/29/2008: Maraxus - Properly restores display depth and frequency.
'***************************************************
    Dim typDevM As typDevMODE
    Dim lRes As Long
    
    If Not bNoResChange Then
    
        lRes = EnumDisplaySettings(0, ENUM_CURRENT_SETTINGS, typDevM)
        
        With typDevM
            .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL Or DM_DISPLAYFREQUENCY
            .dmPelsWidth = oldResWidth
            .dmPelsHeight = oldResHeight
            .dmBitsPerPel = oldDepth
            .dmDisplayFrequency = oldFrequency
        End With
        
        lRes = ChangeDisplaySettings(typDevM, CDS_TEST)
    End If
End Sub
'*****************************
'*****************************

'*****************************
'MOD CARTELES
'*****************************
Sub InitCartel(Ley As String, Grh As Integer)
If Not Cartel Then
    Leyenda = Ley
    textura = Grh
    Cartel = True
    ReDim LeyendaFormateada(0 To (Len(Ley) \ (MAXLONG \ 2)))
                
    Dim i As Integer, k As Integer, anti As Integer
    anti = 1
    k = 0
    i = 0
    Call DarFormato(Leyenda, i, k, anti)
    i = 0
    Do While LeyendaFormateada(i) <> "" And i < UBound(LeyendaFormateada)
        
       i = i + 1
    Loop
    ReDim Preserve LeyendaFormateada(0 To i)
Else
    Exit Sub
End If
End Sub

Private Function DarFormato(s As String, i As Integer, k As Integer, anti As Integer)
If anti + i <= Len(s) + 1 Then
    If ((i >= MAXLONG) And mid$(s, anti + i, 1) = " ") Or (anti + i = Len(s)) Then
        LeyendaFormateada(k) = mid(s, anti, i + 1)
        k = k + 1
        anti = anti + i + 1
        i = 0
    Else
        i = i + 1
    End If
    Call DarFormato(s, i, k, anti)
End If
End Function

Sub DibujarCartel()
If Not Cartel Then Exit Sub
Dim Light(3) As Long
Light(0) = RGB(255, 255, 255)
Light(1) = RGB(255, 255, 255)
Light(2) = RGB(255, 255, 255)
Light(3) = RGB(255, 255, 255)
Dim X As Integer, Y As Integer
X = XPosCartel + 20
Y = YPosCartel + 60
Call DDrawTransGrhIndextoSurface(textura, XPosCartel, YPosCartel, 0, Light)
Dim j As Integer, desp As Integer

For j = 0 To UBound(LeyendaFormateada)
    Fonts_Render_String LeyendaFormateada(j), X, Y + desp, -1
    desp = desp + (frmMain.Font.Size) + 5
Next
End Sub
'*****************************
'*****************************

'*****************************
'MOD PREVINSTANCE
'*****************************
''
' Creates a Named Mutex. Private function, since we will use it just to check if a previous instance of the app is running.
'
' @param mutexName The name of the mutex, should be universally unique for the mutex to be created.

Private Function CreateNamedMutex(ByRef mutexName As String) As Boolean
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/04/07
'Last Modified by: Juan Martín Sotuyo Dodero (Maraxus) - Changed Security Atributes to make it work in all OS
'***************************************************
    Dim SA As SECURITY_ATTRIBUTES
    
    With SA
        .bInheritHandle = 0
        .lpSecurityDescriptor = 0
        .nLength = LenB(SA)
    End With
    
    mutexHID = CreateMutex(SA, False, "Global\" & mutexName)
    
    CreateNamedMutex = Not (Err.LastDllError = ERROR_ALREADY_EXISTS) 'check if the mutex already existed
End Function

''
' Checks if there's another instance of the app running, returns True if there is or False otherwise.

Public Function FindPreviousInstance() As Boolean
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/04/07
'
'***************************************************
    'We try to create a mutex, the name could be anything, but must contain no backslashes.
    If CreateNamedMutex("UniqueNameThatActuallyCouldBeAnything") Then
        'There's no other instance running
        FindPreviousInstance = False
    Else
        'There's another instance running
        FindPreviousInstance = True
    End If
End Function

''
' Closes the client, allowing other instances to be open.

Public Sub ReleaseInstance()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/04/07
'
'***************************************************
    Call ReleaseMutex(mutexHID)
    Call CloseHandle(mutexHID)
End Sub
'*****************************
'*****************************

'*******************************
'MOD APPLICATION
'*******************************
Public Function IsAppActive() As Boolean
'***************************************************
'Author: Juan Martín Sotuyo Dodero (maraxus)
'Last Modify Date: 03/03/2007
'Checks if this is the active application or not
'***************************************************
    IsAppActive = (GetActiveWindow <> 0)
End Function
'******************************
'******************************

'******************************
'MOD AREAS
'******************************

Public Sub CambioDeArea(ByVal X As Byte, ByVal Y As Byte)
    Dim loopX As Long, loopY As Long
    
    MinLimiteX = (X \ 9 - 1) * 9
    MaxLimiteX = MinLimiteX + 26
    
    MinLimiteY = (Y \ 9 - 1) * 9
    MaxLimiteY = MinLimiteY + 26
    
    For loopX = 1 To 100
        For loopY = 1 To 100
            
            If (loopY < MinLimiteY) Or (loopY > MaxLimiteY) Or (loopX < MinLimiteX) Or (loopX > MaxLimiteX) Then
                'Erase NPCs
                
                If MapData(loopX, loopY).CharIndex > 0 Then
                    If MapData(loopX, loopY).CharIndex <> UserCharIndex Then
                        Call EraseChar(MapData(loopX, loopY).CharIndex)
                    End If
                End If
                
                'Erase OBJs
                MapData(loopX, loopY).ObjGrh.GrhIndex = 0
            End If
        Next loopY
    Next loopX
    
    Call RefreshAllChars
End Sub
'*******************************
'*******************************

'*******************************
'MOD COMANDOS
'*******************************
Public Sub Cargar_List()
    Dim X As Byte
    
    COM() = Split(Commands, ",")
    
    For X = 0 To UBound(COM)
        frmMain.LComm.AddItem "/" & COM(X)
    Next X
End Sub

Public Sub SearchCommand(Command As String)
If Opciones.AutoComandos = False Then
    Dim CommandLen, i As Integer
    
    CommandLen = Len(Command)
    
    If CommandLen > 0 Then
    
        With frmMain
        
            For i = 0 To .LComm.ListCount
                If LCase(mid(.LComm.List(i), 1, CommandLen)) = LCase(Command) Then
                    .SendTxt.Text = .LComm.List(i)
                    .SendTxt.SelStart = CommandLen
                    .SendTxt.SelLength = Len(.LComm.List(i)) - CommandLen
                    .LComm.ListIndex = i
                    
                    Exit For
                End If
            Next
        
        End With
        
    End If
End If
End Sub

'***************************
'***************************

'***************************
'MACROS
'***************************
Public Function DoAccionTecla(ByVal Tecla As String)
 On Error Resume Next
    Dim Comando As String
    Comando = GetVar(App.Path & "\Init\Config.cfg", "Macros", Tecla)
     Call ParseUserCommand(Comando)
   
End Function
'***************************
'***************************

Public Sub General_Associate_Icon()
On Error Resume Next

If FileExist(App.Path & "\Recursos\bwaoext.ico", vbNormal) Then
    Dim obj As New cAssociate, System32 As String  'Only Declare if file exist
    Dim Buffer As String * 256
    Dim Tam As Long

        'Get And Set System32 Directory
        Tam = GetSystemDirectory(Buffer, Len(Buffer))
        System32 = Left$(Buffer, Tam)
        
        obj.Extension = ".wao"
        obj.Title = "Recursos WinterAO Ultimate"
        obj.Class = "Clase.wao"
        obj.AppCommand = App.Path & "\WinterAO Ultimate Launcher.exe"
            CopyFile App.Path & "\Recursos\waoext.ico", System32 & "\waoext.ico", True
        obj.DefaultIcon = System32 & "\waoext.ico"
        obj.Associate
    
        Set obj = Nothing
        
        'Kill App.Path & "\Recursos\waoext.ico"
Else
    Exit Sub
End If
End Sub
