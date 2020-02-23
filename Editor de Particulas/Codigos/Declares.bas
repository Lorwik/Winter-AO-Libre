Attribute VB_Name = "Mod_Declaraciones"

Public indexs

Option Explicit
Public font_list() As D3DXFont
Public mode As Boolean
Public temp_rgb(3) As Long
Public LuzGrh(3) As Long
Public AlphaY As Byte
 
Rem Desvanecimientos
Public Type tDesv
    Act As Byte
    AlphaX As Byte
End Type
Public TechoDesv As tDesv

'drag&drop cosas
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const RGN_OR = 2

Public DataChanged As Boolean
Public grh_index As Byte

Public Type tCabecera 'Cabecera de los con
    desc As String * 255
    CRC As Long
    MagicWord As Long
End Type

Public MiCabecera As tCabecera

Public SurfaceDB As clsTexManager

#If SeguridadAlkon Then
Public md5 As New clsMD5
#End If

Public Const MIdi_Inicio As Byte = 6

Public RawServersList As String

Public Type tColor
    r As Byte
    g As Byte
    B As Byte
End Type

Public ColoresPJ(0 To 50) As tColor


Public Type tServerInfo
    Ip As String
    Puerto As Integer
    desc As String
    PassRecPort As Integer
End Type

Public currentMidi As Long

Public ServersLst() As tServerInfo
Public ServersRecibidos As Boolean

Public CurServer As Integer

Public CreandoClan As Boolean
Public ClanName As String
Public Site As String

Public UserCiego As Boolean
Public UserEstupido As Boolean

Public NoRes As Boolean 'no cambiar la resolucion

Public RainBufferIndex As Long
Public FogataBufferIndex As Long

Public Const bCabeza = 1
Public Const bPiernaIzquierda = 2
Public Const bPiernaDerecha = 3
Public Const bBrazoDerecho = 4
Public Const bBrazoIzquierdo = 5
Public Const bTorso = 6

'Timers de GetTickCount
Public Const tAt = 2000
Public Const tUs = 600

Public Const PrimerBodyBarco = 84
Public Const UltimoBodyBarco = 87

Public NumEscudosAnims As Integer

Public ArmasHerrero(0 To 100) As Integer
Public ArmadurasHerrero(0 To 100) As Integer
Public ObjCarpintero(0 To 100) As Integer

Public Versiones(1 To 7) As Integer

Public UsaMacro As Boolean
Public CnTd As Byte
Public SecuenciaMacroHechizos As Byte



'[KEVIN]
Public Const MAX_BANCOINVENTORY_SLOTS = 40
Public UserBancoInventory(1 To MAX_BANCOINVENTORY_SLOTS) As Inventory
'[/KEVIN]


Public Tips() As String * 255
Public Const LoopAdEternum = 999

'Direcciones
Public Enum E_Heading
    NORTH = 1
    EAST = 2
    SOUTH = 3
    WEST = 4
End Enum

'Objetos
Public Const MAX_INVENTORY_OBJS = 10000
Public Const MAX_INVENTORY_SLOTS = 20
Public Const MAX_NPC_INVENTORY_SLOTS = 50
Public Const MAXHECHI = 35


Public Actual As Byte
'Particle Groups
Public TotalStreams As Integer
Public StreamData() As Stream
Public Particula() As Stream
'RGB Type
Public Type RGB
    r As Long
    g As Long
    B As Long
End Type

Public Type Stream
    Name As String
    NumOfParticles As Long
    NumGrhs As Long
    id As Long
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
    angle As Long
    vecx1 As Long
    vecx2 As Long
    vecy1 As Long
    vecy2 As Long
    life1 As Long
    life2 As Long
    friction As Long
    spin As Byte
    spin_speedL As Single
    spin_speedH As Single
    AlphaBlend As Byte
    gravity As Byte
    grav_strength As Long
    bounce_strength As Long
    XMove As Byte
    YMove As Byte
    move_x1 As Long
    move_x2 As Long
    move_y1 As Long
    move_y2 As Long
    grh_list() As Long
    colortint(0 To 3) As RGB
    
    speed As Single
    life_counter As Long
    
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer
    
    Radio As Integer
End Type

Public Const MAXSKILLPOINTS = 100

Public Const FLAGORO = 777

Public Const FOgata = 1521

Public Enum Skills
     Suerte = 1
     Magia = 2
     Robar = 3
     Tacticas = 4
     Armas = 5
     Meditar = 6
     Apuñalar = 7
     Ocultarse = 8
     Supervivencia = 9
     Talar = 10
     Comerciar = 11
     Defensa = 12
     Pesca = 13
     Mineria = 14
     Carpinteria = 15
     Herreria = 16
     Liderazgo = 17 ' NOTA: Solia decir "Curacion"
     Domar = 18
     Proyectiles = 19
     Wresterling = 20
     Navegacion = 21
End Enum

Public Const FundirMetal As Integer = 88

'
' Mensajes
'
' MENSAJE_*  --> Mensajes de texto que se muestran en el cuadro de texto
'

'Inventario
Type Inventory
    OBJIndex As Integer
    Name As String
    grhindex As Integer
    '[Alejo]: tipo de datos ahora es Long
    Amount As Long
    '[/Alejo]
    Equipped As Byte
    Valor As Long
    OBJType As Integer
    Def As Integer
    MaxHit As Integer
    MinHit As Integer
End Type


Type tReputacion 'Fama del usuario
    NobleRep As Long
    BurguesRep As Long
    PlebeRep As Long
    LadronesRep As Long
    BandidoRep As Long
    AsesinoRep As Long
    
    Promedio As Long
End Type

Type tEstadisticasUsu
    CiudadanosMatados As Long
    CriminalesMatados As Long
    UsuariosMatados As Long
    NpcsMatados As Long
    Clase As String
    PenaCarcel As Long
End Type

Public Nombres As Boolean

Public MixedKey As Long

'User status vars
Public UserPassword As String
Public UserName As String

Public Musica As Boolean
Public Sound As Boolean

Public SkillPoints As Integer
Public Alocados As Integer
Public flags() As Integer
Public Oscuridad As Integer
Public logged As Boolean
Public NoPuedeUsar As Boolean

'Barrin 30/9/03
Public UserPuedeRefrescar As Boolean

Public UsingSkill As Integer


Public MD5HushYo As String * 16

Public Enum E_MODO
    Normal = 1
    BorrarPj = 2
    CrearNuevoPj = 3
    Dados = 4
    RecuperarPass = 5
End Enum

Public EstadoLogin As E_MODO
   
Public Enum FxMeditar
'    FXMEDITARCHICO = 4
'    FXMEDITARMEDIANO = 5
'    FXMEDITARGRANDE = 6
'    FXMEDITARXGRANDE = 16
    CHICO = 4
    MEDIANO = 5
    GRANDE = 6
    XGRANDE = 16
End Enum


'Server stuff
Public RequestPosTimer As Integer 'Used in main loop
Public stxtbuffer As String 'Holds temp raw data from server
Public stxtbuffercmsg As String 'Holds temp raw data from server
Public SendNewChar As Boolean 'Used during login
Public Connected As Boolean 'True when connected to server
Public DownloadingMap As Boolean 'Currently downloading a map from server
Public UserMap As Integer

'String contants
Public Const ENDC As String * 1 = vbNullChar    'Endline character for talking with server
Public Const ENDL As String * 2 = vbCrLf        'Holds the Endline character for textboxes

'Control
Public prgRun As Boolean 'When true the program ends

Public IPdelServidor As String
Public PuertoDelServidor As String

'
'********** FUNCIONES API ***********
'

Public Declare Function GetTickCount Lib "kernel32" () As Long

'para escribir y leer variables
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

'Teclado
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Lista de cabezas
Public Type tIndiceCabeza
    Head(1 To 4) As Integer
End Type

Public Type tIndiceCuerpo
    Body(1 To 4) As Integer
    HeadOffsetX As Integer
    HeadOffsetY As Integer
End Type

Public Type tIndiceFx
    Animacion As Integer
    OffsetX As Integer
    OffsetY As Integer
End Type


'Old fashion BitBlt function
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
'Added by Juan Martín Sotuyo Dodero
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long) As Long

