Attribute VB_Name = "General"
Option Explicit

'Contiene info acerca de donde se puede encontrar un grh tamaño y animacion
Public Type GrhData
    SX As Integer
    SY As Integer
    
    FileNum As Long
    
    pixelWidth As Integer
    pixelHeight As Integer
    
    TileWidth As Single
    TileHeight As Single
    
    NumFrames As Integer
    Frames() As Long
    
    Speed As Single
    
    active As Boolean
    MiniMap_color As Long
End Type

'apunta a una estructura grhdata y mantiene la animacion
Public Type Grh
    GrhIndex As Integer
    FrameCounter As Single
    Speed As Single
    Started As Byte
    Loops As Integer
    angle As Single
End Type

'Direcciones
Public Enum E_Heading
    NORTH = 1
    EAST = 2
    SOUTH = 3
    WEST = 4
End Enum

'Posicion en un mapa
Public Type Position
    X As Long
    Y As Long
End Type

'Lista de cabezas
Public Type HeadData
    Head(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type

'Lista de cuerpos
Public Type BodyData
    Walk(E_Heading.NORTH To E_Heading.WEST) As Grh
    HeadOffset As Position
End Type

'Lista de las animaciones de las armas
Type WeaponAnimData
    WeaponWalk(E_Heading.NORTH To E_Heading.WEST) As Grh
        '[ANIM ATAK]
    WeaponAttack As Byte
End Type

'Lista de las animaciones de los escudos
Type ShieldAnimData
    ShieldWalk(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type

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

Public HeadData() As HeadData
Public CascoAnimData() As HeadData
Public BodyData() As BodyData
Public WeaponAnimData() As WeaponAnimData
Public ShieldAnimData() As ShieldAnimData
Public FxData() As tIndiceFx
Public GrhData() As GrhData 'Guarda todos los grh

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long

Private Declare Function TransparentBlt Lib "msimg32" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcsrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal crTransparent As Long) As Long
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long
Private Const COLOR_KEY As Long = &HFF000000

Sub Main()
Dim i As Long

frmcarga.Show
DoEvents

frmcarga.Cargando.Caption = "Cargando Graficos"
DoEvents
Call LoadGrhData
'---------------------------------
frmcarga.Cargando.Caption = "Cargando Cabezas"
DoEvents
Call CargarCabezas
'---------------------------------
frmcarga.Cargando.Caption = "Cargando Cascos"
DoEvents
Call CargarCascos
'---------------------------------
frmcarga.Cargando.Caption = "Cargando Cuerpos"
DoEvents
Call CargarCuerpos
'---------------------------------
frmcarga.Cargando.Caption = "Cargando Armas"
DoEvents
Call CargarAnimArmas
'---------------------------------
frmcarga.Cargando.Caption = "Cargando Escudos"
DoEvents
Call CargarAnimEscudos
'---------------------------------
frmcarga.Cargando.Caption = "Cargando Fx's"
DoEvents
Call CargarFxs

For i = 1 To Numheads
    frmExtra.VisorHead.AddItem (i)
Next i

For i = 1 To NumCascos
    frmExtra.VisorCasco.AddItem (i)
Next i

For i = 1 To NumCuerpos
    frmExtra.VisorCuerpos.AddItem (i)
Next i

For i = 1 To NumWeaponAnims
    frmExtra.VisorArmas.AddItem (i)
Next i

For i = 1 To NumEscudosAnims
    frmExtra.VisorEscudos.AddItem (i)
Next i

For i = 1 To NumFxs
    frmExtra.VisorFX.AddItem (i)
Next i

For i = 1 To grhCount
    If GrhData(i).NumFrames > 1 Then
        frmmain.VisorGrh.AddItem (i & " (Animacion)")
    Else
        frmmain.VisorGrh.AddItem (i)
    End If
Next i

If LoadGrhData = True Then
    Unload frmcarga
    frmmain.Show
Else
    MsgBox "Error al cargar los Init"
    End
End If

End Sub

Public Sub DrawGrhtoHdc(desthdc As Long, ByVal grh_index As Long, ByVal screen_x As Integer, ByVal screen_y As Integer, Optional transparent As Boolean = False)

    On Error Resume Next
    
    Dim file_path As String
    Dim Src_X As Integer
    Dim Src_Y As Integer
    Dim src_width As Integer
    Dim src_height As Integer
    Dim hdcsrc As Long
    Dim MaskDC As Long
    Dim PrevObj As Long
    Dim PrevObj2 As Long
    
    If grh_index <= 0 Then Exit Sub
    
    'If it's animated switch grh_index to first frame
    If GrhData(grh_index).NumFrames <> 1 Then
    grh_index = GrhData(grh_index).Frames(1)
    End If
    
    file_path = App.path & "\GRAFICOS\" & GrhData(grh_index).FileNum & ".bmp"
    
    Src_X = GrhData(grh_index).SX
    Src_Y = GrhData(grh_index).SY
    src_width = GrhData(grh_index).pixelWidth
    src_height = GrhData(grh_index).pixelHeight
    
    hdcsrc = CreateCompatibleDC(desthdc)
    
    PrevObj = SelectObject(hdcsrc, LoadPicture(file_path))
    
    If transparent = False Then
        BitBlt desthdc, screen_x, screen_y, src_width, src_height, hdcsrc, Src_X, Src_Y, vbSrcCopy
    Else
        TransparentBlt desthdc, screen_x, screen_y, src_width, src_height, hdcsrc, Src_X, Src_Y, src_width, src_height, COLOR_KEY
    End If
    
    DeleteDC hdcsrc

End Sub

Function GetVar(file As String, Main As String, var As String) As String
'*****************************************************************
'Gets a Var from a text file
'*****************************************************************

Dim l As Integer
Dim Char As String
Dim sSpaces As String ' This will hold the input that the program will retrieve
Dim szReturn As String ' This will be the defaul value if the string is not found

szReturn = ""

sSpaces = Space(5000) ' This tells the computer how long the longest string can be. If you want, you can change the number 75 to any number you wish


getprivateprofilestring Main, var, szReturn, sSpaces, Len(sSpaces), file

GetVar = RTrim(sSpaces)
GetVar = Left(GetVar, Len(GetVar) - 1)

End Function
Public Function FileExists(ByVal file As String) As Boolean
    FileExists = Dir$(file, vbArchive) <> ""
End Function

Public Function DirExists(ByVal path As String) As Boolean
    DirExists = Dir$(path, vbDirectory) <> ""
End Function
Public Sub WriteVar(file As String, Main As String, var As String, value As String)
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
writeprivateprofilestring Main, var, value, file
End Sub

Public Sub InitGrh(ByRef Grh As Grh, ByVal GrhIndex As Integer, Optional ByVal Started As Byte = 2)
'*****************************************************************
'Sets up a grh. MUST be done before rendering
'*****************************************************************
If GrhIndex = 0 Then Exit Sub
    Grh.GrhIndex = GrhIndex
    
    If Started = 2 Then
        If GrhData(Grh.GrhIndex).NumFrames > 1 Then
            Grh.Started = 1
        Else
            Grh.Started = 0
        End If
    Else
        'Make sure the graphic can be started
        If GrhData(Grh.GrhIndex).NumFrames = 1 Then Started = 0
        Grh.Started = Started
    End If
    
    
    If Grh.Started Then
        Grh.Loops = -1
    Else
        Grh.Loops = 0
    End If
    
    Grh.FrameCounter = 1
    Grh.Speed = GrhData(Grh.GrhIndex).Speed
End Sub
