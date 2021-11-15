Attribute VB_Name = "Mod_WAO"
'=-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-=
'=-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-=Modulo Creado por Lorwik-=-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-=
'=-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-=
'En este modulo esta todos los codigos de los sistemas "mierdas"=-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-=-=
'=-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-=
'LwK Secure AntiClientes editados

'Consola
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_EXSTYLE As Long = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const WS_EX_TRANSPARENT As Long = &H20&
'/Consola

Public Const LwKSecure As String * 18 = "LIK229X1U%REXD-"
'/LwK Secure AntiClientes editados
'***********************************************************
'MOD APIS
    Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    
    Public Const WM_SETTEXT = &HC
    Public Const WM_GETTEXT = &HD
    Public Const WM_GETTEXTLENGTH = &HE
    Public Const EM_SETREADONLY = &HCF
    

Public Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lptypDevMode As Any) As Boolean
Public Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lptypDevMode As Any, ByVal dwFlags As Long) As Long

Public Const CCDEVICENAME = 32
Public Const CCFORMNAME = 32
Public Const DM_BITSPERPEL = &H40000
Public Const DM_PELSWIDTH = &H80000
Public Const DM_PELSHEIGHT = &H100000
Public Const CDS_UPDATEREGISTRY = &H1
Public Const CDS_TEST = &H4
Public Const DISP_CHANGE_SUCCESSFUL = 0
Public Const DISP_CHANGE_RESTART = 1

Type typDevMODE
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
'/MOD APIS
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
'SECU LWK
Private Declare Function GetWindowThreadProcessId Lib "user32" _
(ByVal hWnd As Long, lpdwProcessId As Long) As Long

Private Declare Function OpenProcess Lib "kernel32" (ByVal _
dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
ByVal dwProcessId As Long) As Long

Private Declare Function GetExitCodeProcess Lib "kernel32" _
(ByVal hProcess As Long, lpExitCode As Long) As Long

Private Declare Function TerminateProcess Lib "kernel32" _
(ByVal hProcess As Long, ByVal uExitCode As Long) As Long

'/SECU LWK
'***********************************************************
'LOWIK - MUTEX - ANTIDOBLE CLIENTE
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
'/LORWIK - MUTEX ANDIBLE CLIENTE
'***********************************************************
'Renderizado de Personajes en el Crear
Public MiCuerpo As Integer, MiCabeza As Integer
'**************************************************************
'Control del Volumen:
'Controles del volumen total, y de los canales Rigth y Left
'**************************************************************
Public Const MIXER_SETCONTROLDETAILSF_VALUE = &H0&
Public Const MMSYSERR_NOERROR = 0
Public Const MAXPNAMELEN = 32
Public Const MIXER_LONG_NAME_CHARS = 64
Public Const MIXER_SHORT_NAME_CHARS = 16
Public Const MIXER_GETLINEINFOF_COMPONENTTYPE = &H3&
Public Const MIXER_GETCONTROLDETAILSF_VALUE = &H0&
Public Const MIXER_GETLINECONTROLSF_ONEBYTYPE = &H2&
Public Const MIXERLINE_COMPONENTTYPE_DST_FIRST = &H0&
Public Const MIXERLINE_COMPONENTTYPE_SRC_FIRST = &H1000&
Private Declare Function waveOutSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Private Declare Function waveOutGetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, lpdwVolume As Long) As Long
Const MAX_VALUE As Long = 65535
Public Const MIXERLINE_COMPONENTTYPE_DST_SPEAKERS = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 4)
Public Const MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 3)
Public Const MIXERLINE_COMPONENTTYPE_SRC_LINE = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 2)
Public Const MIXERCONTROL_CT_CLASS_FADER = &H50000000
Public Const MIXERCONTROL_CT_UNITS_UNSIGNED = &H30000
Public Const MIXERCONTROL_CONTROLTYPE_FADER = (MIXERCONTROL_CT_CLASS_FADER Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_VOLUME = (MIXERCONTROL_CONTROLTYPE_FADER + 1)
Public Declare Function mixerClose Lib "winmm.dll" (ByVal hmx As Long) As Long
Public Declare Function mixerGetControlDetails Lib "winmm.dll" Alias "mixerGetControlDetailsA" (ByVal hmxobj As Long, pmxcd As MIXERCONTROLDETAILS, ByVal fdwDetails As Long) As Long
Public Declare Function mixerGetDevCaps Lib "winmm.dll" Alias "mixerGetDevCapsA" (ByVal uMxId As Long, ByVal pmxcaps As MIXERCAPS, ByVal cbmxcaps As Long) As Long
Public Declare Function mixerGetID Lib "winmm.dll" (ByVal hmxobj As Long, pumxID As Long, ByVal fdwId As Long) As Long
Public Declare Function mixerGetLineControls Lib "winmm.dll" Alias "mixerGetLineControlsA" (ByVal hmxobj As Long, pmxlc As MIXERLINECONTROLS, ByVal fdwControls As Long) As Long
Public Declare Function mixerGetLineInfo Lib "winmm.dll" _
               Alias "mixerGetLineInfoA" _
               (ByVal hmxobj As Long, _
               pmxl As MIXERLINE, _
               ByVal fdwInfo As Long) As Long
               
Public Declare Function mixerGetNumDevs Lib "winmm.dll" () As Long
 
Public Declare Function mixerMessage Lib "winmm.dll" _
               (ByVal hmx As Long, _
               ByVal uMsg As Long, _
               ByVal dwParam1 As Long, _
               ByVal dwParam2 As Long) As Long
               
Public Declare Function mixerOpen Lib "winmm.dll" _
               (phmx As Long, _
               ByVal uMxId As Long, _
               ByVal dwCallback As Long, _
               ByVal dwInstance As Long, _
               ByVal fdwOpen As Long) As Long
               
Public Declare Function mixerSetControlDetails Lib "winmm.dll" _
               (ByVal hmxobj As Long, _
               pmxcd As MIXERCONTROLDETAILS, _
               ByVal fdwDetails As Long) As Long
               
Public Declare Sub CopyStructFromPtr Lib "kernel32" _
               Alias "RtlMoveMemory" _
               (struct As Any, _
               ByVal ptr As Long, ByVal CB As Long)
               
Public Declare Sub CopyPtrFromStruct Lib "kernel32" _
               Alias "RtlMoveMemory" _
               (ByVal ptr As Long, _
               struct As Any, _
               ByVal CB As Long)
               
Public Declare Function GlobalAlloc Lib "kernel32" _
               (ByVal wFlags As Long, _
               ByVal dwBytes As Long) As Long
               
Public Declare Function GlobalLock Lib "kernel32" _
               (ByVal hMem As Long) As Long
               
Public Declare Function GlobalFree Lib "kernel32" _
               (ByVal hMem As Long) As Long
 
Public Type MIXERCAPS
    wMid As Integer                   '  manufacturer id
    wPid As Integer                   '  product id
    vDriverVersion As Long            '  version of the driver
    szPname As String * MAXPNAMELEN   '  product name
    fdwSupport As Long                '  misc. support bits
    cDestinations As Long             '  count of destinations
End Type
 
Public Type MIXERCONTROL
    cbStruct As Long           '  size in Byte of MIXERCONTROL
    dwControlID As Long        '  unique control id for mixer device
    dwControlType As Long      '  MIXERCONTROL_CONTROLTYPE_xxx
    fdwControl As Long         '  MIXERCONTROL_CONTROLF_xxx
    cMultipleItems As Long     '  if MIXERCONTROL_CONTROLF_MULTIPLE set
    szShortName As String * MIXER_SHORT_NAME_CHARS  ' short name of control
    szName As String * MIXER_LONG_NAME_CHARS        ' long name of control
    lMinimum As Long           '  Minimum value
    lMaximum As Long           '  Maximum value
    reserved(10) As Long       '  reserved structure space
End Type
 
Public Type MIXERCONTROLDETAILS
    cbStruct As Long       '  size in Byte of MIXERCONTROLDETAILS
    dwControlID As Long    '  control id to get/set details on
    cChannels As Long      '  number of channels in paDetails array
    item As Long           '  hwndOwner or cMultipleItems
    cbDetails As Long      '  size of _one_ details_XX struct
    paDetails As Long      '  pointer to array of details_XX structs
End Type
 
Public Type MIXERCONTROLDETAILS_UNSIGNED
    dwValue As Long        '  value of the control
End Type
 
Public Type MIXERLINE
    cbStruct As Long               '  size of MIXERLINE structure
    dwDestination As Long          '  zero based destination index
    dwSource As Long               '  zero based source index (if source)
    dwLineID As Long               '  unique line id for mixer device
    fdwLine As Long                '  state/information about line
    dwUser As Long                 '  driver specific information
    dwComponentType As Long        '  component type line connects to
    cChannels As Long              '  number of channels line supports
    cConnections As Long           '  number of connections (possible)
    cControls As Long              '  number of controls at this line
    szShortName As String * MIXER_SHORT_NAME_CHARS
    szName As String * MIXER_LONG_NAME_CHARS
    dwType As Long
    dwDeviceID As Long
    wMid  As Integer
    wPid As Integer
    vDriverVersion As Long
    szPname As String * MAXPNAMELEN
End Type
 
Public Type MIXERLINECONTROLS
    cbStruct As Long       '  size in Byte of MIXERLINECONTROLS
    dwLineID As Long       '  line id (from MIXERLINE.dwLineID)
                           '  MIXER_GETLINECONTROLSF_ONEBYID or
    dwControl As Long      '  MIXER_GETLINECONTROLSF_ONEBYTYPE
    cControls As Long      '  count of controls pmxctrl points to
    cbmxctrl As Long       '  size in Byte of _one_ MIXERCONTROL
    pamxctrl As Long       '  pointer to first MIXERCONTROL array
End Type
             
 Public Declare Function waveOutGetNumDevs _
   Lib "winmm.dll" () As Long
 
Public hMixer As Long
Public volCtrl As MIXERCONTROL
Public rc As Long
Public Ok As Boolean
Public VolActual As Long

'Controlar
Public Declare Function SwapMouseButton Lib "user32" ( _
    ByVal bSwap As Long) As Long
 'Old fashion BitBlt function
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
'/PARTE DEL MODULO DE DIBUJADO DE CUENTAS !!!
'***********************************************************

'PARTE DEL MODULO ESTADO MSN PROGRAMABLE
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
 
    Public Type COPYDATASTRUCT
      dwData As Long
      cbData As Long
      lpData As Long
    End Type
 
    Public Const WM_COPYDATA = &H4A
'/PARTE DEL MODULO ESTADO MSN PROGRAMABLE
'***********************************************************
'PARTE DEL MODULO AUTOUPDATE
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal _
cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_TOPMOST = -1
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpfilename As String) As Long
 
Public Darchivo As String
Public strURL As String
Public i As Integer
Public EnProceso As Boolean
'end

'declaraciones zip
Private Type CBChar
    ch(4096) As Byte
End Type
Private Type UNZIPUSERFUNCTION
    UNZIPPrntFunction As Long
    UNZIPSndFunction As Long
    UNZIPReplaceFunction  As Long
    UNZIPPassword As Long
    UNZIPMessage  As Long
    UNZIPService  As Long
    TotalSizeComp As Long
    TotalSize As Long
    CompFactor As Long
    NumFiles As Long
    Comment As Integer
End Type
Private Type UNZIPOPTIONS
    ExtractOnlyNewer  As Long
    SpaceToUnderScore As Long
    PromptToOverwrite As Long
    fQuiet As Long
    ncflag As Long
    ntflag As Long
    nvflag As Long
    nUflag As Long
    nzflag As Long
    ndflag As Long
    noflag As Long
    naflag As Long
    nZIflag As Long
    C_flag As Long
    FPrivilege As Long
    Zip As String
    extractdir As String
End Type
Private Type ZIPnames
    S(0 To 99) As String
End Type
Public Declare Function Wiz_SingleEntryUnzip Lib "unzip32.dll" (ByVal ifnc As Long, ByRef ifnv As ZIPnames, ByVal xfnc As Long, ByRef xfnv As ZIPnames, dcll As UNZIPOPTIONS, Userf As UNZIPUSERFUNCTION) As Long
'end
'/PARTE DEL MODULO AUTOUPDATE
'***********************************************************
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessagges Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const HTCAPTION = 2
Public Const RGN_OR = 2
'***********************************************************
Public Sub HookSurfaceHwnd(frm As Form)
    Call ReleaseCapture
    Call SendMessagges(frm.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub
'***********************************************************
'MINIMAPA
Public Sub DibujarMiniMapa()
On Error Resume Next
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
End Sub
'***********************************************************
Public Sub DibujarMiniMapaUser()
frmMain.UserM.Left = UserPos.X - 1
frmMain.UserM.Top = UserPos.Y - 1
frmMain.UserArea.Left = UserPos.X - 9
frmMain.UserArea.Top = UserPos.Y - 8
End Sub
'***********************************************************
'/MINIMAPA
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
'ANTI CHEAT ENGINE
Public Sub BuscarEngine()
On Error Resume Next
Dim MiObjeto As Object
Set MiObjeto = CreateObject("Wscript.Shell")
Dim X As String
X = "1"
X = MiObjeto.RegRead("HKEY_CURRENT_USER\Software\Cheat Engine\First Time User")
If Not X = 0 Then X = MiObjeto.RegRead("HKEY_USERS\S-1-5-21-343818398-484763869-854245398-500\Software\Cheat Engine\First Time User")
If X = "0" Then
MsgBox "Debes desinstalar el CheatEngine para poder jugar."
End
End If
Set MiObjeto = Nothing
End Sub
'/ANTI CHEAT ENGINE
'***********************************************************
'ORO
Public Function CalculateK(ByVal Gold As Long) As String
'**************************************************************
'Author: lorwik
'Last Modify Date: 11/12/2009
'It calculates the value of the currency in AO
'**************************************************************
    If Val(Gold) < 1000000 Then
        CalculateK = "[" & FormatNumber((Val(Gold) / 1000), 2) & "K]"
    Else
        CalculateK = "[" & FormatNumber((Val(Gold) / 1000000), 2) & "KK]"
    End If
End Function
'/ORO
'***********************************************************
'=-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-=
'=-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-=DIBUJA CUENTAS =-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-=
'=-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-=
Sub DibujaPJ(Surface As DirectDrawSurface7, Grh As Grh, ByVal X As Integer, ByVal Y As Integer, Index As Integer)
On Error Resume Next
Dim r1           As RECT, r2 As RECT, auxr As RECT
Dim iGrhIndex As Integer
If Grh.GrhIndex <= 0 Then Exit Sub
iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
 
With r1
    .Right = GrhData(iGrhIndex).pixelWidth
    .Bottom = GrhData(iGrhIndex).pixelHeight
End With
 
With r2
   .Left = GrhData(iGrhIndex).sX
   .Top = GrhData(iGrhIndex).sY
   .Right = .Left + GrhData(iGrhIndex).pixelWidth
   .Bottom = .Top + GrhData(iGrhIndex).pixelHeight
End With
With auxr
    .Left = 0
  .Top = 0
   .Right = 150
  .Bottom = 150
End With
 
Surface.BltFast X, Y, SurfaceDB.Surface(GrhData(iGrhIndex).FileNum), r2, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
Surface.BltToDC frmCuent.PJ(Index).hDC, auxr, auxr
 
frmCuent.PJ(Index).Refresh
 
End Sub
'***********************************************************
Sub dibujaban(Surface As DirectDrawSurface7, Index As Integer)
 
Dim r2 As RECT, auxr As RECT
 
With r2
   .Left = 0
   .Top = 0
   .Right = 20
   .Bottom = 20
End With
 
With auxr
    .Left = 0
  .Top = 0
   .Right = 150
  .Bottom = 150
End With
 
Surface.SetFontTransparency True
Surface.SetForeColor vbRed
frmCuent.font.Size = 15
Surface.SetFont frmMain.font
'Surface.BltFast x, y, Surface, r2, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
Surface.DrawText 6, 60, "Banned", False
Surface.BltToDC frmCuent.PJ(Index).hDC, auxr, auxr
 
End Sub
'***********************************************************
Sub dibujamuerto(Surface As DirectDrawSurface7, Index As Integer)
 
Dim r2 As RECT, auxr As RECT
 
With r2
   .Left = 0
   .Top = 0
   .Right = 20
   .Bottom = 20
End With
 
With auxr
    .Left = 0
  .Top = 0
   .Right = 150
  .Bottom = 150
End With
 
Surface.SetFontTransparency True
Surface.SetForeColor vbWhite
frmCuent.font.Size = 6
Surface.SetFont frmCuent.font
'Surface.BltFast x, y, Surface, r2, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
Surface.DrawText 5, 10, "MUERTO", False
Surface.BltToDC frmCuent.PJ(Index).hDC, auxr, auxr
 
End Sub
'***********************************************************
Sub DibujarTodo(ByVal Index As Integer, Body As Integer, Head As Integer, Casco As Integer, Shield As Integer, Weapon As Integer, Baned As Integer, Nombre As String, LVL As Integer, Clase As String, muerto As Integer)
 
Dim Grh As Grh
Dim Pos As Integer
Dim loopc As Integer
Dim r As RECT
Dim r2 As RECT
 
Dim YBody As Integer
Dim YYY As Integer
Dim XBody As Integer
Dim BBody As Integer
 
 
With r2
    .Left = 0
  .Top = 0
   .Right = 150
  .Bottom = 150
End With
 
    BackBufferSurface.BltColorFill r, vbBlack
 
If Baned = 1 Then
    Call dibujaban(BackBufferSurface, Index)
End If
 
frmCuent.Nombre(Index).Caption = Nombre
 
frmCuent.Label1(Index).font = frmMain.font
frmCuent.Label1(Index).font = frmMain.font
 
frmCuent.Label1(Index).Caption = LVL
frmCuent.Label2(Index).Caption = Clase
 
XBody = 12
YBody = 15
BBody = 17
 
If muerto = 1 Then
    Body = 8
    Head = 500
    'Arma = 2
    Shield = 2
    Weapon = 2
    XBody = 10
    YBody = 35
    BBody = 16
    Call dibujamuerto(BackBufferSurface, Index)
End If
 
Grh = BodyData(Body).Walk(3)
   
Call DibujaPJ(BackBufferSurface, Grh, XBody, YBody, Index)
 
If muerto = 0 Then YYY = BodyData(Body).HeadOffset.Y
If muerto = 1 Then YYY = -9
 
Pos = YYY + GrhData(GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)).pixelHeight
Grh = HeadData(Head).Head(3)
   
Call DibujaPJ(BackBufferSurface, Grh, BBody, Pos, Index)
   
If Casco <> 2 And Casco > 0 Then
    Grh = CascoAnimData(Casco).Head(3)
    Call DibujaPJ(BackBufferSurface, Grh, BBody, Pos, Index)
End If
 
If Weapon <> 2 And Weapon > 0 Then
    Grh = WeaponAnimData(Weapon).WeaponWalk(3)
    Call DibujaPJ(BackBufferSurface, Grh, XBody, YBody, Index)
End If
 
If Shield <> 2 And Shield > 0 Then
    Grh = ShieldAnimData(Shield).ShieldWalk(3)
    Call DibujaPJ(BackBufferSurface, Grh, XBody, BBody, Index)
End If
   
End Sub
'***********************************************************
'=-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-=
'=-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-=/FIN DE DIBUJADO DE CUENTAS =-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-
'=-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-=
'RELOGEO DE CUENTAS
Public Sub Reloguear()
frmMain.Socket1.HostName = CurServerIp
frmMain.Socket1.RemotePort = CurServerPort
 
If frmMain.Socket1.Connected Then frmMain.Socket1.Disconnect
        If frmConnect.MousePointer = 11 Then Exit Sub
 
        If CheckUserData(False) = True Then
            EstadoLogin = loginaccount
            frmMain.Socket1.Connect
        End If
End Sub
'RELOGEO DE CUENTAS

'**************************************************************
'Control de Voumen:
Function Verificar_tarjeta() As Boolean
     Dim ret As Long
     ret = waveOutGetNumDevs()
     
     If ret >= 0 Then
        Verificar_tarjeta = True
     Else
        Verificar_tarjeta = False
     End If
End Function
'**************************************************************
'Control de Voumen:
Public Function GetVolumeControl(ByRef hMixer As Long, ByVal componentType As Long, ByVal ctrlType As Long, ByRef mxc As MIXERCONTROL) As Boolean
    Dim mxlc As MIXERLINECONTROLS
    Dim mxl As MIXERLINE
    Dim hMem As Long
    mxl.cbStruct = Len(mxl)
    mxl.dwComponentType = componentType
    rc = mixerGetLineInfo(hMixer, mxl, MIXER_GETLINEINFOF_COMPONENTTYPE)
    If (MMSYSERR_NOERROR = rc) Then
        mxlc.cbStruct = Len(mxlc)
        mxlc.dwLineID = mxl.dwLineID
        mxlc.dwControl = ctrlType
        mxlc.cControls = 1
        mxlc.cbmxctrl = Len(mxc)
        hMem = GlobalAlloc(&H40, Len(mxc))
        mxlc.pamxctrl = GlobalLock(hMem)
        mxc.cbStruct = Len(mxc)
        rc = mixerGetLineControls(hMixer, mxlc, MIXER_GETLINECONTROLSF_ONEBYTYPE)
        If (MMSYSERR_NOERROR = rc) Then
            GetVolumeControl = True
            CopyStructFromPtr mxc, mxlc.pamxctrl, Len(mxc)
        Else
            GetVolumeControl = False
        End If
        GlobalFree (hMem)
        Exit Function
    End If
    GetVolumeControl = False
End Function
'**************************************************************
'Control de Voumen:
Public Function GetVolumen(ByRef hMixer As Long, ByRef mxc As MIXERCONTROL) As Long
    Dim mxcd As MIXERCONTROLDETAILS
    Dim Vol As MIXERCONTROLDETAILS_UNSIGNED
    Dim hMem2 As Long
    mxcd.item = 0
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cbStruct = Len(mxcd)
    mxcd.cbDetails = Len(Vol)
    hMem2 = GlobalAlloc(&H40, Len(Vol))
    mxcd.paDetails = GlobalLock(hMem2)
    mxcd.cChannels = 1
    rc = mixerGetControlDetails(hMixer, mxcd, MIXER_GETCONTROLDETAILSF_VALUE)
    CopyStructFromPtr Vol, mxcd.paDetails, Len(Vol)
    GlobalFree (hMem2)
    If (rc = MMSYSERR_NOERROR) Then
        GetVolumen = Vol.dwValue
    Else
        GetVolumen = -1&
    End If
End Function
'**************************************************************
'Control de Voumen:
Public Function SetVolumeControl(ByVal hMixer As Long, mxc As MIXERCONTROL, ByVal Volume As Long) As Boolean
    Dim mxcd As MIXERCONTROLDETAILS
    Dim Vol As MIXERCONTROLDETAILS_UNSIGNED
    Dim hMem As Long
    mxcd.item = 0
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cbStruct = Len(mxcd)
    mxcd.cbDetails = Len(Vol)
    hMem = GlobalAlloc(&H40, Len(Vol))
    mxcd.paDetails = GlobalLock(hMem)
    mxcd.cChannels = 1
    Vol.dwValue = Volume
    CopyPtrFromStruct mxcd.paDetails, Vol, Len(Vol)
    rc = mixerSetControlDetails(hMixer, mxcd, MIXER_SETCONTROLDETAILSF_VALUE)
    GlobalFree (hMem)
    If (MMSYSERR_NOERROR = rc) Then
        SetVolumeControl = True
    Else
        SetVolumeControl = False
    End If
End Function
'**************************************************************
'Control de Voumen:
 Function OpenMixer() As Long
    rc = mixerOpen(hMixer, 0, 0, 0, 0)
    If ((MMSYSERR_NOERROR <> rc)) Then
        OpenMixer = -1
        Exit Function
    End If
    Ok = GetVolumeControl(hMixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, MIXERCONTROL_CONTROLTYPE_VOLUME, volCtrl)
    If (Ok = True) Then
        OpenMixer = GetVolumen(hMixer, volCtrl)
    Else
        OpenMixer = -1
    End If
End Function
'**************************************************************
'Control de Voumen:
Sub CloseMixer()
    On Error Resume Next
    Call mixerClose(hMixer)
End Sub
'**************************************************************
'Control de Voumen:
Public Property Get Volumen() As Byte
Dim v As Long
v = GetVolumen(hMixer, volCtrl)
Volumen = CByte((v / 65535) * 100)
End Property
'**************************************************************
'Control de Voumen:
Public Property Let Volumen(ByVal NewValue As Byte)
    Dim v As Long
    If (NewValue > 100) Then
       MsgBox "El Valor máximo no puede ser superior a 100", vbCritical
       NewValue = 100
    End If
    v = CLng(NewValue * 65535) / 100
    If Not (v > volCtrl.lMaximum Or v < volCtrl.lMinimum) Then
        Call SetVolumeControl(hMixer, volCtrl, v)
    End If
End Property
'**************************************************************
'Control de Voumen:
Sub ActualizaVolumen()
Volumen = frmOpciones.Slider2.value
End Sub
'**************************************************************
'Control de Voumen Rigth/Left:
Sub Get_Balance()
Dim Volumen As Long
Dim Der As Long
Dim Izq As Long
Dim st As Variant
    Call waveOutGetVolume(0, Volumen)
    Der = Volumen And MAX_VALUE
    st = hex(Volumen And -MAX_VALUE)
    If Len(st) > 4 Then
        st = mid(st, 1, Len(st) - 4)
    Else
        st = "0"
    End If
    Izq = CLng("&h" & st)
    frmOpciones.RigthSlider.value = (Der / MAX_VALUE) * 100
    frmOpciones.LeftSlider.value = (Izq / MAX_VALUE) * 100
    frmOpciones.RigthSlider.TickFrequency = 5
    frmOpciones.LeftSlider.TickFrequency = 5
End Sub
'************************************************************
'Control de Voumen Rigth/Left:
Sub Set_Balance()
Dim Der As Long
Dim Izq As Long
Dim Volumen As Long
    Der = (frmOpciones.RigthSlider.value * MAX_VALUE) / 100
    Izq = (frmOpciones.LeftSlider.value * MAX_VALUE) / 100
    Volumen = Val("&h" & hex(Izq) & String(4 - Len(hex(Der)), "0") & hex(Der) & "&")
    Call waveOutSetVolume(0, Volumen)
End Sub
'***********************************************************
'GENERADOR DE CODIGOS !!
Public Function GenerateKey() As String

GenerateKey = RandomNumber(1, 9) & Chr(97 + Rnd() * 862150000 Mod 26) & RandomNumber(1, 9) & Chr(97 + Rnd() * 862150000 Mod 26) & Chr(97 + Rnd() * 862150000 Mod 26) & RandomNumber(1, 9)

End Function
'/GENERADOR DE CODIGOS !!
'**************************************************************
'Renderizado de Personajes en el Crear:
Private Sub DrawGrafico(Grh As Grh, ByVal X As Byte, ByVal Y As Byte)

Dim r2 As RECT, auxr As RECT
Dim iGrhIndex As Integer

    If Grh.GrhIndex <= 0 Then Exit Sub
    
    iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
        
    With r2
        .Left = GrhData(iGrhIndex).sX
        .Top = GrhData(iGrhIndex).sY
        .Right = .Left + GrhData(iGrhIndex).pixelWidth
        .Bottom = .Top + GrhData(iGrhIndex).pixelHeight
    End With
    
    With auxr
        .Left = 0
        .Top = 0
        .Right = 50
        .Bottom = 65
    End With
    
    BackBufferSurface.BltFast X, Y, SurfaceDB.Surface(GrhData(iGrhIndex).FileNum), r2, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    Call BackBufferSurface.BltToDC(frmCrearPersonaje.PlayerView.hDC, auxr, auxr)

End Sub
'**************************************************************
'Renderizado de Personajes en el Crear:
Sub DibujarCPJ(ByVal MyBody As Integer, ByVal MyHead As Integer)

Dim Grh As Grh
Dim Pos As Integer
Dim r2 As RECT

    With r2
        .Left = 0
        .Top = 0
        .Right = 50
        .Bottom = 65
    End With
    
    BackBufferSurface.BltColorFill r2, vbBlack
    
  '  Grh = BodyData(MyBody).Walk(3)
   ' Call DrawGrafico(Grh, 12, 15)
    
    Pos = 0 + 0
    Grh = HeadData(MyHead).Head(3)
    Call DrawGrafico(Grh, 17, Pos)
    
    frmCrearPersonaje.PlayerView.Refresh
    
End Sub
'**************************************************************
'Renderizado de Personajes en el Crear:
Sub DameOpciones()

If frmCrearPersonaje.lstGenero.listIndex < 0 Or frmCrearPersonaje.lstRaza.listIndex < 0 Then
    frmCrearPersonaje.Cabeza.Enabled = False
ElseIf frmCrearPersonaje.lstGenero.listIndex <> -1 And frmCrearPersonaje.lstRaza.listIndex <> -1 Then
    frmCrearPersonaje.Cabeza.Enabled = True
End If

frmCrearPersonaje.Cabeza.Clear
    
Select Case frmCrearPersonaje.lstGenero.List(frmCrearPersonaje.lstGenero.listIndex)
   Case "Hombre"
        Select Case frmCrearPersonaje.lstRaza.List(frmCrearPersonaje.lstRaza.listIndex)
            Case "Humano"
                For i = 1 To 30
                    frmCrearPersonaje.Cabeza.AddItem i
                Next i
                'MiCuerpo = 1
            Case "Elfo"
                For i = 101 To 113
                    If i = 113 Then i = 201
                    frmCrearPersonaje.Cabeza.AddItem i
                Next i
                'MiCuerpo = 2
            Case "Elfo Oscuro"
                For i = 202 To 209
                    frmCrearPersonaje.Cabeza.AddItem i
                Next i
                'MiCuerpo = 3
            Case "Enano"
                For i = 301 To 305
                    frmCrearPersonaje.Cabeza.AddItem i
                Next i
                Case "Orco"
                For i = 511 To 517
                    frmCrearPersonaje.Cabeza.AddItem i
                Next i
                'MiCuerpo = 52
            Case "Gnomo"
                For i = 401 To 406
                    frmCrearPersonaje.Cabeza.AddItem i
                Next i
                'MiCuerpo = 52
            Case Else
                UserHead = 1
                'MiCuerpo = 1
        End Select
   Case "Mujer"
        Select Case frmCrearPersonaje.lstRaza.List(frmCrearPersonaje.lstRaza.listIndex)
            Case "Humano"
                For i = 70 To 76
                    frmCrearPersonaje.Cabeza.AddItem i
                Next i
                'MiCuerpo = 1
            Case "Elfo"
                For i = 170 To 176
                    frmCrearPersonaje.Cabeza.AddItem i
                Next i
                'MiCuerpo = 2
            Case "Elfo Oscuro"
                For i = 270 To 280
                    frmCrearPersonaje.Cabeza.AddItem i
                Next i
                'MiCuerpo = 3
            Case "Gnomo"
                For i = 470 To 474
                    frmCrearPersonaje.Cabeza.AddItem i
                Next i
                'MiCuerpo = 52
                Case "Orco"
                For i = 518 To 522
                    frmCrearPersonaje.Cabeza.AddItem i
                Next i
            Case "Enano"
                UserHead = RandomNumber(1, 3) + 369
                'MiCuerpo = 52
            Case Else
                frmCrearPersonaje.Cabeza.AddItem "70"
                'MiCuerpo = 1
        End Select
End Select

frmCrearPersonaje.PlayerView.Cls

End Sub
'ANTIDOBLE LCIENTE
'***********************************************************
Private Function CreateNamedMutex(ByRef mutexName As String) As Boolean
    Dim sa As SECURITY_ATTRIBUTES
    
    With sa
        .bInheritHandle = 0
        .lpSecurityDescriptor = 0
        .nLength = LenB(sa)
    End With
    
    mutexHID = CreateMutex(sa, False, "Global\" & mutexName)
    
    CreateNamedMutex = Not (Err.LastDllError = ERROR_ALREADY_EXISTS) 'check if the mutex already existed
End Function
'***********************************************************
Public Function FindPreviousInstance() As Boolean
    'We try to create a mutex, the name could be anything, but must contain no backslashes.
    If CreateNamedMutex("UniqueNameThatActuallyCouldBeAnything") Then
        'There's no other instance running
        FindPreviousInstance = False
    Else
        'There's another instance running
        FindPreviousInstance = True
    End If
End Function
'***********************************************************
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
'***********************************************************
'/ANTIDOBLE CLIENTE
'MACROS
Public Function DoAccionTecla(ByVal Tecla As String)
 On Error Resume Next
Dim Accion As Byte
    Accion = GetVar(IniPath & "Macros.bin", Tecla, "Accion")
   
    If Accion = 1 Then
        Dim Comando As String
        Comando = GetVar(IniPath & "Macros.bin", Tecla, "Comando")
            Call SendData("/" & Comando)
    ElseIf Accion = 2 Then
        Dim Usar As Byte
        Usar = GetVar(IniPath & "Macros.bin", Tecla, "UsarItem")
            Call SendData("USA" & Usar)
    ElseIf Accion = 3 Then
        Dim Equipar As Byte
        Equipar = GetVar(IniPath & "Macros.bin", Tecla, "EquiparItem")
            Call SendData("EQUI" & Equipar)
    ElseIf Accion = 4 Then
        Dim Hechizo As Byte
        Hechizo = GetVar(IniPath & "Macros.bin", Tecla, "LanzarHechizo")
            Call SendData("HL" & Hechizo)
            Call SendData("KU" & Magia)
    ElseIf Accion <> 1 Or 2 Or 3 Or 4 Then
        Exit Function
    ElseIf Accion = "" Then
        Exit Function
    End If
   
End Function
'***********************************************************
Public Function DibujarMacros(ByVal Tecla As Integer)
 
Dim Accion As Byte
    Accion = GetVar(IniPath & "Macros.bin", "F" & Tecla, "Accion")
 
    If Accion = 1 Then
    Call Extract_File2(Graphics, App.Path & "\ARCHIVOS", "2427.bmp", Windows_Temp_Dir, False)    ' Windows_Temp_Dir, False)
        frmMain.Macros(Tecla).Picture = LoadPicture(Windows_Temp_Dir & "2427.bmp")
        Delete_File (Windows_Temp_Dir & "2427.bmp")
       
    ElseIf Accion = 2 Then
        Dim Usar As Byte
            Usar = GetVar(IniPath & "Macros.bin", "F" & Tecla, "UsarItem")
        Dim Grh As Integer
            Grh = Inventario.GrhIndex(Usar)
                Call DibujarMacrosUsarEquipar(Grh, Tecla)
         
    ElseIf Accion = 3 Then
        Dim Equipar As Byte
            Equipar = GetVar(IniPath & "Macros.bin", "F" & Tecla, "EquiparItem")
        Dim Grhs As Integer
            Grhs = Inventario.GrhIndex(Equipar)
                Call DibujarMacrosUsarEquipar(Grhs, Tecla)
               
    ElseIf Accion = 4 Then
    
        Call Extract_File2(Graphics, App.Path & "\ARCHIVOS", "617.bmp", Windows_Temp_Dir, False)    ' Windows_Temp_Dir, False)
    
        frmMain.Macros(Tecla).Picture = LoadPicture(Windows_Temp_Dir & "617.bmp")
        Delete_File (Windows_Temp_Dir & "617.bmp")
      
    ElseIf Accion <> 1 Or 2 Or 3 Or 4 Then
        Exit Function
    End If
End Function
'***********************************************************
Public Function DibujarMacrosUsarEquipar(ByVal Grh As Integer, ByVal Tecla As Integer)
Dim SR As RECT, DR As RECT
SR.Left = 0
SR.Top = 0
SR.Right = 34
SR.Bottom = 34
DR.Left = 0
DR.Top = 0
DR.Right = 34
DR.Bottom = 34
Call DrawGrhtoHdc(frmMain.Macros(Tecla).hWnd, frmMain.Macros(Tecla).hDC, Grh, SR, DR)
End Function
'***********************************************************
Public Function CargarMacros()
    Dim i As Byte
        For i = 1 To 12
            Call DibujarMacros(i)
        Next i
End Function
'/MACROS
'***********************************************************
'ScreenShoots
Public Sub Capturar_Guardar(Path As String)
Clipboard.Clear
keybd_event VK_SNAPSHOT, 1, 0, 0
DoEvents
    If Clipboard.GetFormat(vbCFBitmap) Then
            SavePicture Clipboard.GetData(vbCFBitmap), Path
    Else
        Call AddtoRichTextBox(frmMain.RecTxt, "Error al tomar la foto", 255, 255, 255, False, False, False)
    End If

End Sub
'***********************************************************
' SECU LWK
Public Sub CerrarProceso(TítuloVentana As String)
Dim hProceso As Long
Dim lEstado As Long
Dim idProc As Long
Dim winHwnd As Long

winHwnd = FindWindow(vbNullString, TítuloVentana)
If winHwnd = 0 Then
Debug.Print "El proceso no está abierto": Exit Sub
End If
Call GetWindowThreadProcessId(winHwnd, idProc)

' Obtenemos el handle al proceso
hProceso = OpenProcess(PROCESS_TERMINATE Or _
PROCESS_QUERY_INFORMATION, 0, idProc)
If hProceso <> 0 Then
' Comprobamos estado del proceso
GetExitCodeProcess hProceso, lEstado
If lEstado = STILL_ACTIVE Then
' Cerramos el proceso
If TerminateProcess(hProceso, 9) <> 0 Then
Debug.Print "Proceso cerrado"
Else
Debug.Print "No se pudo matar el proceso"
End If
End If
' Cerramos el handle asociado al proceso
CloseHandle hProceso
Else
Debug.Print "No se pudo tener acceso al proceso"
End If
End Sub
'***********************************************************
Public Sub SecuLwK()
'Lorwik> La verdad esque no me gusta mucho, pero mejor esto que nada xD
If FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.1.1")) Then
    Call HayExterno("CHEAT ENGINE 5.1.1")
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.0")) Then
    Call HayExterno("CHEAT ENGINE 5.0")
ElseIf FindWindow(vbNullString, UCase$("Pts")) Then
    Call HayExterno("Auto Pots")
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.2")) Then
    Call HayExterno("CHEAT ENGINE 5.2")
ElseIf FindWindow(vbNullString, UCase$("SOLOCOVO?")) Then
    Call HayExterno("SOLOCOVO?")
ElseIf FindWindow(vbNullString, UCase$("-=[ANUBYS RADAR]=-")) Then
    Call HayExterno("-=[ANUBYS RADAR]=-")
ElseIf FindWindow(vbNullString, UCase$("CRAZY SPEEDER 1.05")) Then
    Call HayExterno("CRAZY SPEEDER 1.05")
ElseIf FindWindow(vbNullString, UCase$("SET !XSPEED.NET")) Then
    Call HayExterno("SET !XSPEED.NET")
ElseIf FindWindow(vbNullString, UCase$("SPEEDERXP V1.80 - UNREGISTERED")) Then
    Call HayExterno("SPEEDERXP V1.80 - UNREGISTERED")
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.3")) Then
    Call HayExterno("CHEAT ENGINE 5.3")
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.1")) Then
    Call HayExterno("CHEAT ENGINE 5.1")
ElseIf FindWindow(vbNullString, UCase$("A SPEEDER")) Then
    Call HayExterno("A SPEEDER")
ElseIf FindWindow(vbNullString, UCase$("MEMO :P")) Then
    Call HayExterno("MEMO :P")
ElseIf FindWindow(vbNullString, UCase$("ORK4M VERSION 1.5")) Then
    Call HayExterno("ORK4M VERSION 1.5")
ElseIf FindWindow(vbNullString, UCase$("BY FEDEX")) Then
    Call HayExterno("By Fedex")
ElseIf FindWindow(vbNullString, UCase$("!XSPEED.NET +4.59")) Then
    Call HayExterno("!Xspeeder")
ElseIf FindWindow(vbNullString, UCase$("CAMBIA TITULOS DE CHEATS BY FEDEX")) Then
    Call HayExterno("Cambia titulos")
ElseIf FindWindow(vbNullString, UCase$("NEWENG OCULTO")) Then
    Call HayExterno("Cambia titulos")
ElseIf FindWindow(vbNullString, UCase$("SERBIO ENGINE")) Then
    Call HayExterno("Serbio Engine")
ElseIf FindWindow(vbNullString, UCase$("REYMIX ENGINE 5.3 PUBLIC")) Then
    Call HayExterno("ReyMix Engine")
ElseIf FindWindow(vbNullString, UCase$("REY ENGINE 5.2")) Then
    Call HayExterno("ReyMix Engine")
ElseIf FindWindow(vbNullString, UCase$("AUTOCLICK - BY NIO_SHOOTER")) Then
    Call HayExterno("AutoClick")
ElseIf FindWindow(vbNullString, UCase$("TONNER MINER! :D [REG][SKLOV] 2.0")) Then
    Call HayExterno("Tonner")
ElseIf FindWindow(vbNullString, UCase$("Buffy The vamp Slayer")) Then
    Call HayExterno("Buffy The vamp Slayer")
ElseIf FindWindow(vbNullString, UCase$("Blorb Slayer 1.12.552 (BETA)")) Then
    Call HayExterno("Blorb Slayer 1.12.552 (BETA)")
ElseIf FindWindow(vbNullString, UCase$("PumaEngine3.0")) Then
    Call HayExterno("PumaEngine3.0")
ElseIf FindWindow(vbNullString, UCase$("Vicious Engine 5.0")) Then
    Call HayExterno("Vicious Engine 5.0")
ElseIf FindWindow(vbNullString, UCase$("AkumaEngine33")) Then
    Call HayExterno("AkumaEngine33")
ElseIf FindWindow(vbNullString, UCase$("Spuc3ngine")) Then
    Call HayExterno("Spuc3ngine")
ElseIf FindWindow(vbNullString, UCase$("Ultra Engine")) Then
    Call HayExterno("Ultra Engine")
ElseIf FindWindow(vbNullString, UCase$("Engine")) Then
    Call HayExterno("Engine")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V5.4")) Then
    Call HayExterno("Cheat Engine V5.4")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V4.4")) Then
    Call HayExterno("Cheat Engine V4.4")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V4.4 German Add-On")) Then
    Call HayExterno("Cheat Engine V4.4 German Add-On")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V4.3")) Then
    Call HayExterno("Cheat Engine V4.3")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V4.2")) Then
    Call HayExterno("Cheat Engine V4.2")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V4.1.1")) Then
    Call HayExterno("Cheat Engine V4.1.1")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V3.3")) Then
    Call HayExterno("Cheat Engine V3.3")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V3.2")) Then
    Call HayExterno("Cheat Engine V3.2")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V3.1")) Then
    Call HayExterno("Cheat Engine V3.1")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine")) Then
    Call HayExterno("Cheat Engine")
ElseIf FindWindow(vbNullString, UCase$("danza engine 5.2.150")) Then
    Call HayExterno("danza engine 5.2.150")
ElseIf FindWindow(vbNullString, UCase$("zenx engine")) Then
    Call HayExterno("zenx engine")
ElseIf FindWindow(vbNullString, UCase$("MACROMAKER")) Then
    Call HayExterno("Macro Maker")
ElseIf FindWindow(vbNullString, UCase$("MACREOMAKER - EDIT MACRO")) Then
    Call HayExterno("Macro Maker")
ElseIf FindWindow(vbNullString, UCase$("By Fedex")) Then
    Call HayExterno("Macro Fedex")
ElseIf FindWindow(vbNullString, UCase$("Macro Mage 1.0")) Then
    Call HayExterno("Macro Mage")
ElseIf FindWindow(vbNullString, UCase$("Auto* v0.4 (c) 2001 [Agresión] Powa")) Then
    Call HayExterno("Macro Fisher")
ElseIf FindWindow(vbNullString, UCase$("Kizsada")) Then
    Call HayExterno("Macro K33")
ElseIf FindWindow(vbNullString, UCase$("Makro K33")) Then
    Call HayExterno("Macro K33")
ElseIf FindWindow(vbNullString, UCase$("Super Saiyan")) Then
    Call HayExterno("El Chit del Geri")
ElseIf FindWindow(vbNullString, UCase$("Makro-Piringulete")) Then
    Call HayExterno("Piringulete")
ElseIf FindWindow(vbNullString, UCase$("Makro-Piringulete 2003")) Then
    Call HayExterno("Piringulete 2003")
ElseIf FindWindow(vbNullString, UCase$("TUKY2005")) Then
    Call HayExterno("Makro Tuky")
ElseIf FindWindow(vbNullString, UCase$("Macro Configurable")) Then
Call HayExterno("Macro Configurable")


End If

End Sub
'***********************************************************
'/SECU LWK
Function UnEncryptStr(ByVal S As String, ByVal p As String) As String
Dim i As Integer, r As String
Dim C1 As Integer, C2 As Integer

r = ""
If Len(p) > 0 Then
    For i = 1 To Len(S)
        C1 = Asc(mid(S, i, 1))
        If i > Len(p) Then
            C2 = Asc(mid(p, i Mod Len(p) + 1, 1))
        Else
            C2 = Asc(mid(p, i, 1))
        End If
        C1 = C1 - C2 - 64
        If Sgn(C1) = -1 Then C1 = 256 + C1
        r = r + Chr(C1)
    Next i
Else
    r = S
End If

UnEncryptStr = r
End Function
'***********************************************************
Public Sub ListarServidores()
On Error Resume Next
frmConnect.LstServidores.Clear
    Dim Cargar As String
    Dim i As Byte
    Cargar = frmConnect.Inet1.OpenURL("http://www.updatewinterao.com.ar/server/sinfo.php")
    For i = 1 To 8
        Servidor(i).IP = ReadField(1, ReadField(i, Cargar, Asc("|")), Asc(":"))
        Servidor(i).Puerto = ReadField(2, ReadField(i, Cargar, Asc("|")), Asc(":"))
        Servidor(i).Nombre = ReadField(3, ReadField(i, Cargar, Asc("|")), Asc(":"))
    Next i
    'Primero dame IP de sv asi te la encripto, ok espera 1 seg
    
    For i = 1 To 10
        frmConnect.LstServidores.AddItem Servidor(i).Nombre
    Next i
    
    frmConnect.Text1.Text = Servidor(1).IP
    frmConnect.Text2.Text = Servidor(1).Puerto
End Sub
'***********************************************************
'Transparencia consola

Public Sub Make_Transparent_Richtext(ByVal hWnd As Long)

    Call SetWindowLong(hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)

End Sub
'***********************************************************
'Logs
Public Sub LogError(desc As String)
On Error Resume Next
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\errores.log" For Append As #nfile
Print #nfile, desc
Close #nfile
End Sub
'***********************************************************
Public Sub LogCustom(desc As String)
On Error Resume Next
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\custom.log" For Append As #nfile
Print #nfile, Now & " " & desc
Close #nfile
End Sub
'/Logs
'***********************************************************
'SUB MAIN
Sub CerrarCliente()
EngineRun = False
    frmCargando.Show
    LiberarObjetosDX

    If Not bNoResChange Then
        Dim typDevM As typDevMODE
        Dim lRes As Long
        
        lRes = EnumDisplaySettings(0, 0, typDevM)
        With typDevM
            .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
            .dmPelsWidth = oldResWidth
            .dmPelsHeight = oldResHeight
        End With
        lRes = ChangeDisplaySettings(typDevM, CDS_TEST)
    End If

    'Destruimos los objetos públicos creados
    Set SurfaceDB = Nothing
    Set Dialogos = Nothing
    Set DialogosClanes = Nothing
    Set Audio = Nothing
    Set Inventario = Nothing
    Call UnloadAllForms
    
    
    'Lorwik> Borramos la basura del pc del usuario.
     Delete_File (Windows_Temp_Dir & "2427.bmp")
     Delete_File (Windows_Temp_Dir & "617.bmp")
    
End
End Sub
'***********************************************************
Sub IniciarEngine()

 Set SurfaceDB = New clsSurfaceManDyn
 IniciarObjetosDirectX
 Call InitTileEngine(frmMain.hWnd, frmMain.MainViewShp.Top - 1, frmMain.MainViewShp.Left + 3, 32, 32, Round(frmMain.MainViewShp.Height / 32), Round(frmMain.MainViewShp.Width / 32), 9)

        'Inicializamos el sonido
    Musica = Not ClientSetup.bNoMusic
    Sound = Not ClientSetup.bNoSound
 Call Audio.Initialize(DirectX, frmMain.hWnd, Windows_Temp_Dir & "\", Windows_Temp_Dir & "\")

    'Inicializamos el inventario gráfico
    Call Inventario.Initialize(frmMain.picInv)
End Sub
'***********************************************************
Sub IniciarMP3()
        'Comprobacion del MP3
     If GetVar(App.Path & "\Init\config.ini", "INIT", "MP3") = 0 Then
        MPTres = False
        Else
        MPTres = True
     End If
     If MPTres Then
          Call Extract_File2(wav, App.Path & "\ARCHIVOS", "click.wav", Windows_Temp_Dir, False)
        
        Set MP3P = New clsMP3Player
           Call Extract_File2(mp3, App.Path & "\ARCHIVOS\", "1.mp3", Windows_Temp_Dir, False)
        MP3P.stopMP3
        MP3P.mp3file = Windows_Temp_Dir & "1.mp3"
        MP3P.playMP3
        MP3P.Volume = 1000
        Delete_File (Windows_Temp_Dir & "1.mp3")
     End If
End Sub
'***********************************************************
Sub IniciarCliente()
    frmMain.Socket1.Startup
    Call CargarAnimsExtra
    Call frmCargando.progresoConDelay(75)
    Call CargarAnimArmas
    Call CargarAnimEscudos
    Call CargarVersiones
    Call CargarColores
    Call frmCargando.progresoConDelay(80)
End Sub
'/SUB MAIN
