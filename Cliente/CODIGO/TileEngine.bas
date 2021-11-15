Attribute VB_Name = "Mod_TileEngine"

Option Explicit

Rem Mannakia Fps Libres
Public timerElapsedTime As Single
Public timerTicksPerFrame As Double

'    C       O       N       S      T
'Map sizes in tiles
Public Const XMaxMapSize = 100
Public Const XMinMapSize = 1
Public Const YMaxMapSize = 100
Public Const YMinMapSize = 1

Public Const GrhFogata = 1521

'bltbit constant
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source

'    T       I      P      O      S
'Encabezado bmp
Type BITMAPFILEHEADER
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type

'Info del encabezado del bmp
Type BITMAPINFOHEADER
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

'Posicion en un mapa
Public Type Position
    X As Long
    Y As Long
End Type

'Posicion en el Mundo
Public Type WorldPos
    Map As Integer
    X As Integer
    Y As Integer
End Type

'Contiene info acerca de donde se puede encontrar un grh
'tamaño y animacion
Public Type GrhData
    sX As Integer
    sY As Integer
    FileNum As Integer
    pixelWidth As Integer
    pixelHeight As Integer
    TileWidth As Single
    TileHeight As Single
    NumFrames As Integer
    Frames(1 To 25) As Integer
    Speed As Integer
    Active As Boolean
    MiniMap_color As Long
End Type

'apunta a una estructura grhdata y mantiene la animacion
Public Type Grh
    GrhIndex As Integer
    FrameCounter As Single
    Speed As Single
    Started As Byte
End Type

'Lista de cuerpos
Public Type BodyData
    Walk(1 To 4) As Grh
    HeadOffset As Position
End Type

'Lista de cabezas
Public Type HeadData
    Head(1 To 4) As Grh
End Type

'Lista de las animaciones de las armas
Type WeaponAnimData
    WeaponWalk(1 To 4) As Grh
    '[ANIM ATAK]
    WeaponAttack As Byte
End Type

'Lista de las animaciones de los escudos
Type ShieldAnimData
    ShieldWalk(1 To 4) As Grh
End Type


'Lista de cuerpos
Public Type FxData
    Fx As Grh
    OffsetX As Long
    OffsetY As Long
End Type

'Apariencia del personaje
Public Type Char

    Aura_Index As Integer
    Aura As Grh
    
    Active As Byte
    Heading As Byte ' As E_Heading ?
    Pos As Position
    
    iHead As Integer
    iBody As Integer
    Body As BodyData
    Head As HeadData
    Casco As HeadData
    Arma As WeaponAnimData
    Escudo As ShieldAnimData
    UsandoArma As Boolean
    Fx As Integer
    FxLoopTimes As Integer
    Criminal As Byte
    
    Nombre As String
    
    scrollDirectionX As Integer
    scrollDirectionY As Integer
    
    Moving As Byte
    MoveOffsetX As Single
    MoveOffsetY As Single
    ServerIndex As Integer
    
    pie As Boolean
    muerto As Boolean
    invisible As Boolean
    priv As Byte
    
End Type

'Info de un objeto
Public Type Obj
    OBJIndex As Integer
    Amount As Integer
End Type

'Tipo de las celdas del mapa
Public Type MapBlock
    Graphic(1 To 4) As Grh
    CharIndex As Integer
    ObjGrh As Grh
    ObjName As String
    
    NPCIndex As Integer
    OBJInfo As Obj
    TileExit As WorldPos
    Blocked As Byte
    
    Trigger As Integer
End Type

'Info de cada mapa
Public Type MapInfo
    Music As String
    Name As String
    StartPos As WorldPos
    MapVersion As Integer
    
    'ME Only
    Changed As Byte
End Type


Public IniPath As String
Public MapPath As String


'Bordes del mapa
Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

'Status del user
Public CurMap As Integer 'Mapa actual
Public UserIndex As Integer
Public UserMoving As Byte
Public UserBody As Integer
Public UserHead As Integer
Public UserPos As Position 'Posicion
Public AddtoUserPos As Position 'Si se mueve
Public UserCharIndex As Integer

Public UserMaxAGU As Integer
Public UserMinAGU As Integer
Public UserMaxHAM As Integer
Public UserMinHAM As Integer
Public UserGuerra As Boolean 'Guerras

Public UserGLDBOV As Long
Public UserBOVItem As Long

Public EngineRun As Boolean
Public FramesPerSec As Integer
Public FramesPerSecCounter As Long

'Tamaño del la vista en Tiles
Public WindowTileWidth As Integer
Public WindowTileHeight As Integer

'Offset del desde 0,0 del main view
Public MainViewTop As Integer
Public MainViewLeft As Integer

'Cuantos tiles el engine mete en el BUFFER cuando
'dibuja el mapa. Ojo un tamaño muy grande puede
'volver el engine muy lento
Public TileBufferSize As Integer

'Handle to where all the drawing is going to take place
Public DisplayFormhWnd As Long

'Tamaño de los tiles en pixels
Public TilePixelHeight As Integer
Public TilePixelWidth As Integer

'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Totales?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

Public NumBodies As Integer
Public Numheads As Integer
Public NumFxs As Integer

Public NumChars As Integer
Public LastChar As Integer
Public NumWeaponAnims As Integer
Public NumShieldAnims As Integer

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Graficos¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

Public lastTime As Long 'Para controlar la velocidad


'[CODE]:MatuX'
Public MainDestRect   As RECT
'[END]'
Public MainViewRect   As RECT
Public BackBufferRect As RECT

Public MainViewWidth As Integer
Public MainViewHeight As Integer




'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Graficos¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public GrhData() As GrhData 'Guarda todos los grh
Public BodyData() As BodyData
Public HeadData() As HeadData
Public FxData() As FxData
Public WeaponAnimData() As WeaponAnimData
Public ShieldAnimData() As ShieldAnimData
Public CascoAnimData() As HeadData
Public Grh() As Grh 'Animaciones publicas
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Mapa?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public MapData() As MapBlock ' Mapa
Public MapInfo As MapInfo ' Info acerca del mapa en uso
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Usuarios?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'
'epa ;)
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿API?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'Blt
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?


'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'       [CODE 000]: MatuX
'
Public bRain        As Boolean 'está raineando?
Public bTecho       As Boolean 'hay techo?
Public brstTick     As Long

Private iFrameIndex As Byte  'Frame actual de la LL
Private llTick      As Long  'Contador

Public charlist(1 To 10000) As Char

'estados internos del surface (read only)
Public Enum TextureStatus
    tsOriginal = 0
    tsNight = 1
    tsFog = 2
End Enum


#If ConAlfaB Then

Private Declare Function BltAlphaFast Lib "vbabdx" (ByRef lpDDSDest As Any, ByRef lpDDSSource As Any, ByVal iWidth As Long, ByVal iHeight As Long, _
        ByVal pitchSrc As Long, ByVal pitchDst As Long, ByVal dwMode As Long) As Long
Private Declare Function BltEfectoNoche Lib "vbabdx" (ByRef lpDDSDest As Any, ByVal iWidth As Long, ByVal iHeight As Long, _
        ByVal pitchDst As Long, ByVal dwMode As Long) As Long

'LORWIK - DECLARACIONES DE LA LIBRERIA DLL vbadb
Public Declare Function vbDABLalphablend16 Lib "vbDABL" (ByVal iMode As Integer, ByVal bColorKey As Integer, _
ByRef sPtr As Any, ByRef dPtr As Any, ByVal iAlphaVal As Integer, ByVal iWidth As Integer, ByVal iHeight As Integer, _
ByVal isPitch As Integer, ByVal idPitch As Integer, ByVal iColorKey As Integer) As Integer
Public Declare Function vbDABLcolorblend16555 Lib "vbDABL" (ByRef sPtr As Any, ByRef dPtr As Any, ByVal alpha_val%, _
ByVal Width%, ByVal Height%, ByVal sPitch%, ByVal dPitch%, ByVal rVal%, ByVal gVal%, ByVal bVal%) As Long
Public Declare Function vbDABLcolorblend16565 Lib "vbDABL" (ByRef sPtr As Any, ByRef dPtr As Any, ByVal alpha_val%, _
ByVal Width%, ByVal Height%, ByVal sPitch%, ByVal dPitch%, ByVal rVal%, ByVal gVal%, ByVal bVal%) As Long
Public Declare Function vbDABLcolorblend16555ck Lib "vbDABL" (ByRef sPtr As Any, ByRef dPtr As Any, ByVal alpha_val%, _
ByVal Width%, ByVal Height%, ByVal sPitch%, ByVal dPitch%, ByVal rVal%, ByVal gVal%, ByVal bVal%) As Long
Public Declare Function vbDABLcolorblend16565ck Lib "vbDABL" (ByRef sPtr As Any, ByRef dPtr As Any, ByVal alpha_val%, _
ByVal Width%, ByVal Height%, ByVal sPitch%, ByVal dPitch%, ByVal rVal%, ByVal gVal%, ByVal bVal%) As Long
'/LORWIK - DECLARACIONES DE LA LIBRERIA DLL vbadb

#End If

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Sub CargarCabezas()
On Error Resume Next
Dim N As Integer, i As Integer, Numheads As Integer, Index As Integer

Dim Miscabezas() As tIndiceCabeza

N = FreeFile
Open App.Path & "\init\Cabezas.ind" For Binary Access Read As #N

'cabecera
Get #N, , MiCabecera

'num de cabezas
Get #N, , Numheads

'Resize array
ReDim HeadData(0 To Numheads + 1) As HeadData
ReDim Miscabezas(0 To Numheads + 1) As tIndiceCabeza

For i = 1 To Numheads
    Get #N, , Miscabezas(i)
    InitGrh HeadData(i).Head(1), Miscabezas(i).Head(1), 0
    InitGrh HeadData(i).Head(2), Miscabezas(i).Head(2), 0
    InitGrh HeadData(i).Head(3), Miscabezas(i).Head(3), 0
    InitGrh HeadData(i).Head(4), Miscabezas(i).Head(4), 0
Next i

Close #N

End Sub

Sub CargarCascos()
On Error Resume Next
Dim N As Integer, i As Integer, NumCascos As Integer, Index As Integer

Dim Miscabezas() As tIndiceCabeza

N = FreeFile
Open App.Path & "\init\Cascos.ind" For Binary Access Read As #N

'cabecera
Get #N, , MiCabecera

'num de cabezas
Get #N, , NumCascos

'Resize array
ReDim CascoAnimData(0 To NumCascos + 1) As HeadData
ReDim Miscabezas(0 To NumCascos + 1) As tIndiceCabeza

For i = 1 To NumCascos
    Get #N, , Miscabezas(i)
    InitGrh CascoAnimData(i).Head(1), Miscabezas(i).Head(1), 0
    InitGrh CascoAnimData(i).Head(2), Miscabezas(i).Head(2), 0
    InitGrh CascoAnimData(i).Head(3), Miscabezas(i).Head(3), 0
    InitGrh CascoAnimData(i).Head(4), Miscabezas(i).Head(4), 0
Next i

Close #N

End Sub

Sub CargarCuerpos()
On Error Resume Next
Dim N As Integer, i As Integer
Dim NumCuerpos As Integer
Dim MisCuerpos() As tIndiceCuerpo

N = FreeFile
Open App.Path & "\init\Personajes.ind" For Binary Access Read As #N

'cabecera
Get #N, , MiCabecera

'num de cabezas
Get #N, , NumCuerpos

'Resize array
ReDim BodyData(0 To NumCuerpos + 1) As BodyData
ReDim MisCuerpos(0 To NumCuerpos + 1) As tIndiceCuerpo

For i = 1 To NumCuerpos
    Get #N, , MisCuerpos(i)
    InitGrh BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0
    InitGrh BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0
    InitGrh BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0
    InitGrh BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0
    BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
    BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY
Next i

Close #N

End Sub
Sub CargarFxs()
On Error Resume Next
Dim N As Integer, i As Integer
Dim NumFxs As Integer
Dim MisFxs() As tIndiceFx

N = FreeFile
Open App.Path & "\init\Fxs.ind" For Binary Access Read As #N

'cabecera
Get #N, , MiCabecera

'num de cabezas
Get #N, , NumFxs

'Resize array
ReDim FxData(0 To NumFxs + 1) As FxData
ReDim MisFxs(0 To NumFxs + 1) As tIndiceFx

For i = 1 To NumFxs
    Get #N, , MisFxs(i)
    Call InitGrh(FxData(i).Fx, MisFxs(i).Animacion, 1)
    FxData(i).OffsetX = MisFxs(i).OffsetX
    FxData(i).OffsetY = MisFxs(i).OffsetY
Next i

Close #N

End Sub

Sub CargarTips()
On Error Resume Next
Dim N As Integer, i As Integer
Dim NumTips As Integer

N = FreeFile
Open App.Path & "\init\Tips.ayu" For Binary Access Read As #N

'cabecera
Get #N, , MiCabecera

'num de cabezas
Get #N, , NumTips

'Resize array
ReDim Tips(1 To NumTips) As String * 255

For i = 1 To NumTips
    Get #N, , Tips(i)
Next i

Close #N

End Sub
Sub ConvertCPtoTP(StartPixelLeft As Integer, StartPixelTop As Integer, ByVal cx As Single, ByVal cy As Single, tX As Integer, tY As Integer)
'******************************************
'Converts where the user clicks in the main window
'to a tile position
'******************************************
Dim HWindowX As Integer
Dim HWindowY As Integer

cx = cx - StartPixelLeft
cy = cy - StartPixelTop

HWindowX = (WindowTileWidth \ 2)
HWindowY = (WindowTileHeight \ 2)

'Figure out X and Y tiles
cx = (cx \ TilePixelWidth)
cy = (cy \ TilePixelHeight)

If cx > HWindowX Then
    cx = (cx - HWindowX)

Else
    If cx < HWindowX Then
        cx = (0 - (HWindowX - cx))
    Else
        cx = 0
    End If
End If

If cy > HWindowY Then
    cy = (0 - (HWindowY - cy))
Else
    If cy < HWindowY Then
        cy = (cy - HWindowY)
    Else
        cy = 0
    End If
End If

tX = UserPos.X + cx
tY = UserPos.Y + cy

End Sub






Sub MakeChar(ByVal CharIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal X As Integer, ByVal Y As Integer, ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer)

On Error Resume Next

'Apuntamos al ultimo Char
If CharIndex > LastChar Then LastChar = CharIndex

NumChars = NumChars + 1

If Arma = 0 Then Arma = 2
If Escudo = 0 Then Escudo = 2
If Casco = 0 Then Casco = 2

charlist(CharIndex).iHead = Head
charlist(CharIndex).iBody = Body
charlist(CharIndex).Head = HeadData(Head)
charlist(CharIndex).Body = BodyData(Body)
charlist(CharIndex).Arma = WeaponAnimData(Arma)
'[ANIM ATAK]
charlist(CharIndex).Arma.WeaponAttack = 0

charlist(CharIndex).Escudo = ShieldAnimData(Escudo)
charlist(CharIndex).Casco = CascoAnimData(Casco)

charlist(CharIndex).Heading = Heading

'Reset moving stats
charlist(CharIndex).Moving = 0
charlist(CharIndex).MoveOffsetX = 0
charlist(CharIndex).MoveOffsetY = 0

'Update position
charlist(CharIndex).Pos.X = X
charlist(CharIndex).Pos.Y = Y

'Make active
charlist(CharIndex).Active = 1

'Plot on map
MapData(X, Y).CharIndex = CharIndex

End Sub

Sub ResetCharInfo(ByVal CharIndex As Integer)

    charlist(CharIndex).Active = 0
    charlist(CharIndex).Criminal = 0
    charlist(CharIndex).Fx = 0
    charlist(CharIndex).FxLoopTimes = 0
    charlist(CharIndex).invisible = False



    charlist(CharIndex).Moving = 0
    charlist(CharIndex).muerto = False
    charlist(CharIndex).Nombre = ""
    charlist(CharIndex).pie = False
    charlist(CharIndex).Pos.X = 0
    charlist(CharIndex).Pos.Y = 0
    charlist(CharIndex).UsandoArma = False

End Sub


Sub EraseChar(ByVal CharIndex As Integer)
On Error Resume Next

'*****************************************************************
'Erases a character from CharList and map
'*****************************************************************

charlist(CharIndex).Active = 0

'Update lastchar
If CharIndex = LastChar Then
    Do Until charlist(LastChar).Active = 1
        LastChar = LastChar - 1
        If LastChar = 0 Then Exit Do
    Loop
End If


MapData(charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y).CharIndex = 0

Call ResetCharInfo(CharIndex)

'Update NumChars
NumChars = NumChars - 1

End Sub

Sub InitGrh(ByRef Grh As Grh, ByVal GrhIndex As Integer, Optional Started As Byte = 2)
'*****************************************************************
'Sets up a grh. MUST be done before rendering
'*****************************************************************

Grh.GrhIndex = GrhIndex

If Started = 2 Then
    If GrhData(Grh.GrhIndex).NumFrames > 1 Then
        Grh.Started = 1
    Else
        Grh.Started = 0
    End If
Else
    Grh.Started = Started
End If

Grh.FrameCounter = 1

'Moficamos la velocidad para FPS Libres
Grh.Speed = GrhData(Grh.GrhIndex).Speed


End Sub

Sub MoveCharbyHead(ByVal CharIndex As Integer, ByVal nHeading As E_Heading)
'*****************************************************************
'Starts the movement of a character in nHeading direction
'*****************************************************************
Dim addX As Integer
Dim addY As Integer
Dim X As Integer
Dim Y As Integer
Dim nX As Integer
Dim nY As Integer

X = charlist(CharIndex).Pos.X
Y = charlist(CharIndex).Pos.Y

'Figure out which way to move
Select Case nHeading

    Case E_Heading.NORTH
        addY = -1

    Case E_Heading.EAST
        addX = 1

    Case E_Heading.SOUTH
        addY = 1
    
    Case E_Heading.WEST
        addX = -1
        
End Select

nX = X + addX
nY = Y + addY

MapData(nX, nY).CharIndex = CharIndex
charlist(CharIndex).Pos.X = nX
charlist(CharIndex).Pos.Y = nY
MapData(X, Y).CharIndex = 0

charlist(CharIndex).MoveOffsetX = -1 * (TilePixelWidth * addX)
charlist(CharIndex).MoveOffsetY = -1 * (TilePixelHeight * addY)

charlist(CharIndex).scrollDirectionX = addX
charlist(CharIndex).scrollDirectionY = addY

charlist(CharIndex).Moving = 1
charlist(CharIndex).Heading = nHeading

If UserEstado <> 1 Then Call DoPasosFx(CharIndex)

'areas viejos
If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
    Debug.Print UserCharIndex
    Call EraseChar(CharIndex)
End If

End Sub

Public Sub DoFogataFx()
    Dim location As Position
    
    If bFogata Then
        bFogata = HayFogata(location)
        If Not bFogata Then
            Call Audio.StopWave(FogataBufferIndex)
            FogataBufferIndex = 0
        End If
    Else
        bFogata = HayFogata(location)
        If bFogata And FogataBufferIndex = 0 Then FogataBufferIndex = Audio.PlayWave("fuego.wav", location.X, location.Y, LoopStyle.Enabled)
    End If
End Sub

Function EstaPCarea(ByVal Index2 As Integer) As Boolean

Dim X As Integer, Y As Integer

For Y = UserPos.Y - MinYBorder + 1 To UserPos.Y + MinYBorder - 1
  For X = UserPos.X - MinXBorder + 1 To UserPos.X + MinXBorder - 1
            
            If MapData(X, Y).CharIndex = Index2 Then
                EstaPCarea = True
                Exit Function
            End If
        
  Next X
Next Y

EstaPCarea = False

End Function
Public Function TickON(Cual As Integer, Cont As Integer) As Boolean
Static tickCount(200) As Integer
If Cont = 999 Then Exit Function
tickCount(Cual) = tickCount(Cual) + 1
If tickCount(Cual) < Cont Then
    TickON = False
Else
    tickCount(Cual) = 0
    TickON = True
End If
End Function
Sub DoPasosFx(ByVal CharIndex As Integer)

 If Not UserNavegando Then
        With charlist(CharIndex)
            If Not .muerto And EstaPCarea(CharIndex) Then
                .pie = Not .pie
        'Lorwik> Efectos de sonido de caminata segun donde este
   If UserEquitando = True And CharIndex = UserCharIndex And EstaPCarea(CharIndex) Then
        If TickON(0, 3) Then Call Audio.PlayWave(SND_GALOPE, .Pos.X, .Pos.Y)
        
                ElseIf MapData(.Pos.X, .Pos.Y).Graphic(1).GrhIndex >= 6000 And MapData(.Pos.X, .Pos.Y).Graphic(1).GrhIndex <= 6559 Then
                    If .pie Then
                        Call Audio.PlayWave("23.wav", .Pos.X, .Pos.Y)
                    Else
                        Call Audio.PlayWave("24.wav", .Pos.X, .Pos.Y)
                    End If
                    
                ElseIf MapData(.Pos.X, .Pos.Y).Graphic(1).GrhIndex >= 17986 And MapData(.Pos.X, .Pos.Y).Graphic(1).GrhIndex <= 18604 Then
                 If .pie Then
                        Call Audio.PlayWave("x23.wav", .Pos.X, .Pos.Y)
                    Else
                        Call Audio.PlayWave("x24.wav", .Pos.X, .Pos.Y)
                    End If
                    
                    Else
                    If .pie Then
                        Call Audio.PlayWave("241.wav", .Pos.X, .Pos.Y)
                    Else
                        Call Audio.PlayWave("244.wav", .Pos.X, .Pos.Y)
                    End If
                End If
            End If
    
    
    If UserNavegando Then
' TODO : Actually we would have to check if the CharIndex char is in the water or not....
        Call Audio.PlayWave(SND_NAVEGANDO, .Pos.X, .Pos.Y)
    End If
    
    End With
    End If
    End Sub


Sub MoveCharbyPos(ByVal CharIndex As Integer, ByVal nX As Integer, ByVal nY As Integer)

On Error Resume Next

Dim X As Integer
Dim Y As Integer
Dim addX As Integer
Dim addY As Integer
Dim nHeading As E_Heading



X = charlist(CharIndex).Pos.X
Y = charlist(CharIndex).Pos.Y

MapData(X, Y).CharIndex = 0

addX = nX - X
addY = nY - Y

If Sgn(addX) = 1 Then
    nHeading = E_Heading.EAST
End If

If Sgn(addX) = -1 Then
    nHeading = E_Heading.WEST
End If

If Sgn(addY) = -1 Then
    nHeading = E_Heading.NORTH
End If

If Sgn(addY) = 1 Then
    nHeading = E_Heading.SOUTH
End If

MapData(nX, nY).CharIndex = CharIndex


charlist(CharIndex).Pos.X = nX
charlist(CharIndex).Pos.Y = nY

charlist(CharIndex).MoveOffsetX = -1 * (TilePixelWidth * addX)
charlist(CharIndex).MoveOffsetY = -1 * (TilePixelHeight * addY)
charlist(CharIndex).scrollDirectionX = Sgn(addX)
charlist(CharIndex).scrollDirectionY = Sgn(addY)


charlist(CharIndex).Moving = 1
charlist(CharIndex).Heading = nHeading

'parche para que no medite cuando camina
Dim fxCh As Integer
fxCh = charlist(CharIndex).Fx
If fxCh = FxMeditar.CHICO Or fxCh = FxMeditar.GRANDE Or fxCh = FxMeditar.MEDIANO Or fxCh = FxMeditar.XGRANDE Then
    charlist(CharIndex).Fx = 0
    charlist(CharIndex).FxLoopTimes = 0
End If

If Not EstaPCarea(CharIndex) Then Dialogos.QuitarDialogo (CharIndex)

If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
    Call EraseChar(CharIndex)
End If

End Sub

Sub MoveScreen(ByVal nHeading As E_Heading)
'******************************************
'Starts the screen moving in a direction
'******************************************
Dim X As Integer
Dim Y As Integer
Dim tX As Integer
Dim tY As Integer

'Figure out which way to move
Select Case nHeading

    Case E_Heading.NORTH
        Y = -1

    Case E_Heading.EAST
        X = 1

    Case E_Heading.SOUTH
        Y = 1
    
    Case E_Heading.WEST
        X = -1
        
End Select

'Fill temp pos
tX = UserPos.X + X
tY = UserPos.Y + Y

'Check to see if its out of bounds
If tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder Then
    Exit Sub
Else
    'Start moving... MainLoop does the rest
    AddtoUserPos.X = X
    UserPos.X = tX
    AddtoUserPos.Y = Y
    UserPos.Y = tY
    UserMoving = 1
   
End If


    

End Sub
Private Function HayFogata(ByRef location As Position) As Boolean
    Dim j As Long
    Dim k As Long
    
    For j = UserPos.X - 8 To UserPos.X + 8
        For k = UserPos.Y - 6 To UserPos.Y + 6
            If InMapBounds(j, k) Then
                If MapData(j, k).ObjGrh.GrhIndex = GrhFogata Then
                    location.X = j
                    location.Y = k
                    
                    HayFogata = True
                    Exit Function
                End If
            End If
        Next k
    Next j
End Function

Function NextOpenChar() As Integer
'*****************************************************************
'Finds next open char slot in CharList
'*****************************************************************
Dim loopc As Integer
Dim Dale As Boolean

loopc = 1
Do While charlist(loopc).Active And Dale
    loopc = loopc + 1
    Dale = (loopc <= UBound(charlist))
Loop

NextOpenChar = loopc

End Function


Sub LoadGrhData()
'*****************************************************************
'Loads Grh.dat
'*****************************************************************

On Error GoTo ErrorHandler

Dim Grh As Integer
Dim Frame As Integer
Dim tempint As Integer

'Resize arrays
ReDim GrhData(1 To 32000) As GrhData

'Open files
Open IniPath & "Graficos.ind" For Binary Access Read As #1
Seek #1, 1

Get #1, , MiCabecera
Get #1, , tempint
Get #1, , tempint
Get #1, , tempint
Get #1, , tempint
Get #1, , tempint

'Fill Grh List

'Get first Grh Number
Get #1, , Grh

Do Until Grh <= 0
        GrhData(Grh).Active = True
    'Get number of frames
    Get #1, , GrhData(Grh).NumFrames
    If GrhData(Grh).NumFrames <= 0 Then GoTo ErrorHandler
    
    If GrhData(Grh).NumFrames > 1 Then
    
        'Read a animation GRH set
        For Frame = 1 To GrhData(Grh).NumFrames
        
            Get #1, , GrhData(Grh).Frames(Frame)
            If GrhData(Grh).Frames(Frame) <= 0 Or GrhData(Grh).Frames(Frame) > 32000 Then
                GoTo ErrorHandler
            End If
        
        Next Frame
    
        Get #1, , GrhData(Grh).Speed
        GrhData(Grh).Speed = GrhData(Grh).NumFrames / 0.018
        If GrhData(Grh).Speed <= 0 Then GoTo ErrorHandler
        
        'Compute width and height
        GrhData(Grh).pixelHeight = GrhData(GrhData(Grh).Frames(1)).pixelHeight
        If GrhData(Grh).pixelHeight <= 0 Then GoTo ErrorHandler
        
        GrhData(Grh).pixelWidth = GrhData(GrhData(Grh).Frames(1)).pixelWidth
        If GrhData(Grh).pixelWidth <= 0 Then GoTo ErrorHandler
        
        GrhData(Grh).TileWidth = GrhData(GrhData(Grh).Frames(1)).TileWidth
        If GrhData(Grh).TileWidth <= 0 Then GoTo ErrorHandler
        
        GrhData(Grh).TileHeight = GrhData(GrhData(Grh).Frames(1)).TileHeight
        If GrhData(Grh).TileHeight <= 0 Then GoTo ErrorHandler
    
    Else
    
        'Read in normal GRH data
        Get #1, , GrhData(Grh).FileNum
        If GrhData(Grh).FileNum <= 0 Then GoTo ErrorHandler
        
        Get #1, , GrhData(Grh).sX
        If GrhData(Grh).sX < 0 Then GoTo ErrorHandler
        
        Get #1, , GrhData(Grh).sY
        If GrhData(Grh).sY < 0 Then GoTo ErrorHandler
            
        Get #1, , GrhData(Grh).pixelWidth
        If GrhData(Grh).pixelWidth <= 0 Then GoTo ErrorHandler
        
        Get #1, , GrhData(Grh).pixelHeight
        If GrhData(Grh).pixelHeight <= 0 Then GoTo ErrorHandler
        
        'Compute width and height
        GrhData(Grh).TileWidth = GrhData(Grh).pixelWidth / TilePixelHeight
        GrhData(Grh).TileHeight = GrhData(Grh).pixelHeight / TilePixelWidth
        
        GrhData(Grh).Frames(1) = Grh
            
    End If

    'Get Next Grh Number
    Get #1, , Grh

Loop
'************************************************

Close #1
Dim Count As Long
 
Open IniPath & "minimap.dat" For Binary As #1
    Seek #1, 1
    For Count = 1 To 32000
        If GrhData(Count).Active Then
            Get #1, , GrhData(Count).MiniMap_color
        End If
    Next Count
Close #1
Exit Sub

ErrorHandler:
Close #1
MsgBox "Error while loading the Grh.dat! Stopped at GRH number: " & Grh

End Sub
Function LegalPos(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is legal
'*****************************************************************

'Limites del mapa
If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
    LegalPos = False
    Exit Function
End If

    'Tile Bloqueado?
    If MapData(X, Y).Blocked = 1 Then
        LegalPos = False
        Exit Function
    End If
    
    '¿Hay un personaje?
    If MapData(X, Y).CharIndex > 0 Then
        LegalPos = False
        Exit Function
    End If
   
    If Not UserNavegando Then
        If HayAgua(X, Y) Then
            LegalPos = False
            Exit Function
        End If
    Else
        If Not HayAgua(X, Y) Then
            LegalPos = False
            Exit Function
        End If
    End If
            If UserEquitando Then
            If bTecho = True Then
                LegalPos = False
                Exit Function
            End If
            End If
        
LegalPos = True

End Function




Function InMapLegalBounds(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps
'LEGAL/Walkable bounds
'*****************************************************************

If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
    InMapLegalBounds = False
    Exit Function
End If

InMapLegalBounds = True

End Function

Function InMapBounds(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************

If X < XMinMapSize Or X > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
    InMapBounds = False
    Exit Function
End If

InMapBounds = True

End Function

Sub DDrawGrhtoSurface(Surface As DirectDrawSurface7, Grh As Grh, ByVal X As Integer, ByVal Y As Integer, center As Byte, Animate As Byte)

Dim CurrentGrh As Grh
Dim destRect As RECT
Dim SourceRect As RECT
Dim SurfaceDesc As DDSURFACEDESC2

If Animate Then
    If Grh.Started = 1 Then
        Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed)
        If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
            Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
        End If
    End If
End If


'Figure out what frame to draw (always 1 if not animated)
CurrentGrh.GrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
'Center Grh over X,Y pos
If center Then
    If GrhData(CurrentGrh.GrhIndex).TileWidth <> 1 Then
        X = X - Int(GrhData(CurrentGrh.GrhIndex).TileWidth * 16) + 16 'hard coded for speed
    End If
    If GrhData(CurrentGrh.GrhIndex).TileHeight <> 1 Then
        Y = Y - Int(GrhData(CurrentGrh.GrhIndex).TileHeight * 32) + 32 'hard coded for speed
    End If
End If
With SourceRect
        .Left = GrhData(CurrentGrh.GrhIndex).sX
        .Top = GrhData(CurrentGrh.GrhIndex).sY
        .Right = .Left + GrhData(CurrentGrh.GrhIndex).pixelWidth
        .Bottom = .Top + GrhData(CurrentGrh.GrhIndex).pixelHeight
End With
Surface.BltFast X, Y, SurfaceDB(GrhData(CurrentGrh.GrhIndex).FileNum), SourceRect, DDBLTFAST_WAIT
End Sub

Sub DDrawTransGrhIndextoSurface(Surface As DirectDrawSurface7, Grh As Integer, ByVal X As Integer, ByVal Y As Integer, center As Byte, Animate As Byte)
Dim CurrentGrh As Grh
Dim destRect As RECT
Dim SourceRect As RECT
Dim SurfaceDesc As DDSURFACEDESC2

With destRect
    .Left = X
    .Top = Y
    .Right = .Left + GrhData(Grh).pixelWidth
    .Bottom = .Top + GrhData(Grh).pixelHeight
End With

Surface.GetSurfaceDesc SurfaceDesc

'Draw
If destRect.Left >= 0 And destRect.Top >= 0 And destRect.Right <= SurfaceDesc.lWidth And destRect.Bottom <= SurfaceDesc.lHeight Then
    With SourceRect
        .Left = GrhData(Grh).sX
        .Top = GrhData(Grh).sY
        .Right = .Left + GrhData(Grh).pixelWidth
        .Bottom = .Top + GrhData(Grh).pixelHeight
    End With
    
    Surface.BltFast destRect.Left, destRect.Top, SurfaceDB.Surface(GrhData(Grh).FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
End If

End Sub

'Sub DDrawTransGrhtoSurface(surface As DirectDrawSurface7, Grh As Grh, X As Integer, Y As Integer, Center As Byte, Animate As Byte, Optional ByVal KillAnim As Integer = 0)
'[CODE 000]:MatuX
    Sub DDrawTransGrhtoSurface(Surface As DirectDrawSurface7, Grh As Grh, ByVal X As Integer, ByVal Y As Integer, center As Byte, Animate As Byte, Optional ByVal KillAnim As Integer = 0)
'[END]'
'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
'[CODE]:MatuX
'
'  CurrentGrh.GrhIndex = iGrhIndex
'
'[END]

'Dim CurrentGrh As Grh
Dim iGrhIndex As Integer
'Dim destRect As RECT
Dim SourceRect As RECT
'Dim SurfaceDesc As DDSURFACEDESC2
Dim QuitarAnimacion As Boolean

If Animate Then
    If Grh.Started = 1 Then
        Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed)
        If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
            Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
            If KillAnim Then
                If charlist(KillAnim).FxLoopTimes <> LoopAdEternum Then
                    If charlist(KillAnim).FxLoopTimes > 0 Then charlist(KillAnim).FxLoopTimes = charlist(KillAnim).FxLoopTimes - 1
                    If charlist(KillAnim).FxLoopTimes < 1 Then 'Matamos la anim del fx ;))
                        charlist(KillAnim).Fx = 0
                        Exit Sub
                    End If
                End If
            End If
        End If
    End If
End If

If Grh.GrhIndex = 0 Then Exit Sub

'Figure out what frame to draw (always 1 if not animated)
iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)

'Center Grh over X,Y pos
If center Then
    If GrhData(iGrhIndex).TileWidth <> 1 Then
        X = X - Int(GrhData(iGrhIndex).TileWidth * 16) + 16 'hard coded for speed
    End If
    If GrhData(iGrhIndex).TileHeight <> 1 Then
        Y = Y - Int(GrhData(iGrhIndex).TileHeight * 32) + 32 'hard coded for speed
    End If
End If

With SourceRect
    .Left = GrhData(iGrhIndex).sX
    .Top = GrhData(iGrhIndex).sY
    .Right = .Left + GrhData(iGrhIndex).pixelWidth
    .Bottom = .Top + GrhData(iGrhIndex).pixelHeight
End With


Surface.BltFast X, Y, SurfaceDB.Surface(GrhData(iGrhIndex).FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT

End Sub

#If ConAlfaB = 1 Then
    Sub DDrawTransGrhtoSurfaceAlpha(Surface As DirectDrawSurface7, Grh As Grh, ByVal X As Integer, ByVal Y As Integer, center As Byte, Animate As Byte, Optional ByVal KillAnim As Integer = 0)
'[END]'
'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
'[CODE]:MatuX
'
'  CurrentGrh.GrhIndex = iGrhIndex
'
'[END]

'Dim CurrentGrh As Grh
Dim iGrhIndex As Integer
'Dim destRect As RECT
Dim SourceRect As RECT
'Dim SurfaceDesc As DDSURFACEDESC2
Dim QuitarAnimacion As Boolean


If Animate Then
    If Grh.Started = 1 Then
        Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed)
        If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
            Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
            If KillAnim Then
                If charlist(KillAnim).FxLoopTimes <> LoopAdEternum Then
                    If charlist(KillAnim).FxLoopTimes > 0 Then charlist(KillAnim).FxLoopTimes = charlist(KillAnim).FxLoopTimes - 1
                    If charlist(KillAnim).FxLoopTimes < 1 Then 'Matamos la anim del fx ;))
                        charlist(KillAnim).Fx = 0
                        Exit Sub
                    End If
                End If
            End If
        End If
    End If
End If


If Grh.GrhIndex = 0 Then Exit Sub

'Figure out what frame to draw (always 1 if not animated)
iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)

'Center Grh over X,Y pos
If center Then
    If GrhData(iGrhIndex).TileWidth <> 1 Then
        X = X - Int(GrhData(iGrhIndex).TileWidth * 16) + 16 'hard coded for speed
    End If
    If GrhData(iGrhIndex).TileHeight <> 1 Then
        Y = Y - Int(GrhData(iGrhIndex).TileHeight * 32) + 32 'hard coded for speed
    End If
End If

With SourceRect
    .Left = GrhData(iGrhIndex).sX + IIf(X < 0, Abs(X), 0)
    .Top = GrhData(iGrhIndex).sY + IIf(Y < 0, Abs(Y), 0)
    .Right = .Left + GrhData(iGrhIndex).pixelWidth
    .Bottom = .Top + GrhData(iGrhIndex).pixelHeight
End With

'surface.BltFast X, Y, SurfaceDB.surface(GrhData(iGrhIndex).FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT

Dim src As DirectDrawSurface7
Dim rDest As RECT
Dim dArray() As Byte, sArray() As Byte
Dim ddsdSrc As DDSURFACEDESC2, ddsdDest As DDSURFACEDESC2
Dim Modo As Long

Set src = SurfaceDB.Surface(GrhData(iGrhIndex).FileNum)

src.GetSurfaceDesc ddsdSrc
Surface.GetSurfaceDesc ddsdDest

With rDest
    .Left = X
    .Top = Y
    .Right = X + GrhData(iGrhIndex).pixelWidth
    .Bottom = Y + GrhData(iGrhIndex).pixelHeight
    
    If .Right > ddsdDest.lWidth Then
        .Right = ddsdDest.lWidth
    End If
    If .Bottom > ddsdDest.lHeight Then
        .Bottom = ddsdDest.lHeight
    End If
End With

' 0 -> 16 bits 555
' 1 -> 16 bits 565
' 2 -> 16 bits raro (Sin implementar)
' 3 -> 24 bits
' 4 -> 32 bits

If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0 And ddsdSrc.ddpfPixelFormat.lGBitMask = &H3E0 Then
    Modo = 0
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 And ddsdSrc.ddpfPixelFormat.lGBitMask = &H7E0 Then
    Modo = 1
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 And ddsdSrc.ddpfPixelFormat.lGBitMask = &H7E0 Then
    Modo = 3
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = 65280 And ddsdSrc.ddpfPixelFormat.lGBitMask = 65280 Then
    Modo = 4
Else
    'Modo = 2 '16 bits raro ?
    Surface.BltFast X, Y, src, SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    Exit Sub
End If

Dim SrcLock As Boolean, DstLock As Boolean
SrcLock = False: DstLock = False

On Local Error GoTo HayErrorAlpha

src.Lock SourceRect, ddsdSrc, DDLOCK_WAIT, 0
SrcLock = True
Surface.Lock rDest, ddsdDest, DDLOCK_WAIT, 0
DstLock = True

Surface.GetLockedArray dArray()
src.GetLockedArray sArray()

Call BltAlphaFast(ByVal VarPtr(dArray(X + X, Y)), ByVal VarPtr(sArray(SourceRect.Left * 2, SourceRect.Top)), rDest.Right - rDest.Left, rDest.Bottom - rDest.Top, ddsdSrc.lPitch, ddsdDest.lPitch, Modo)

Surface.Unlock rDest
DstLock = False
src.Unlock SourceRect
SrcLock = False


Exit Sub

HayErrorAlpha:
If SrcLock Then src.Unlock SourceRect
If DstLock Then Surface.Unlock rDest

End Sub
#End If 'ConAlfaB = 1

Sub DrawBackBufferSurface()
    PrimarySurface.Blt MainViewRect, BackBufferSurface, MainDestRect, DDBLT_WAIT
End Sub

Function GetBitmapDimensions(BmpFile As String, ByRef bmWidth As Long, ByRef bmHeight As Long)
'*****************************************************************
'Gets the dimensions of a bmp
'*****************************************************************
Dim BMHeader As BITMAPFILEHEADER
Dim BINFOHeader As BITMAPINFOHEADER

Open BmpFile For Binary Access Read As #1
Get #1, , BMHeader
Get #1, , BINFOHeader
Close #1
bmWidth = BINFOHeader.biWidth
bmHeight = BINFOHeader.biHeight
End Function

Sub DrawGrhtoHdc(hWnd As Long, hDC As Long, Grh As Integer, SourceRect As RECT, destRect As RECT)
    If Grh <= 0 Then Exit Sub
    
    SecundaryClipper.SetHWnd hWnd
    SurfaceDB.Surface(GrhData(Grh).FileNum).BltToDC hDC, SourceRect, destRect
End Sub

Sub RenderScreen(tilex As Integer, tiley As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
On Error Resume Next


If UserCiego Then Exit Sub

Dim Y        As Integer 'Keeps track of where on map we are
Dim X        As Integer 'Keeps track of where on map we are
Dim minY     As Integer 'Start Y pos on current map
Dim maxY     As Integer 'End Y pos on current map
Dim minX     As Integer 'Start X pos on current map
Dim maxX     As Integer 'End X pos on current map
Dim ScreenX  As Integer 'Keeps track of where to place tile on screen
Dim ScreenY  As Integer 'Keeps track of where to place tile on screen
Dim Moved    As Byte
Dim Grh      As Grh     'Temp Grh for show tile and blocked
Dim tempChar As Char
Dim TextX    As Integer
Dim TextY    As Integer
Dim iPPx     As Integer 'Usado en el Layer de Chars
Dim iPPy     As Integer 'Usado en el Layer de Chars
Dim rSourceRect      As RECT    'Usado en el Layer 1
Dim iGrhIndex        As Integer 'Usado en el Layer 1
Dim PixelOffsetXTemp As Integer 'For centering grhs
Dim PixelOffsetYTemp As Integer 'For centering grhs
Dim nX As Integer
Dim nY As Integer

'Figure out Ends and Starts of screen
' Hardcodeado para speed!
minY = (tiley - 15)
maxY = (tiley + 15)
minX = (tilex - 17)
maxX = (tilex + 17)


'Draw floor layer
ScreenY = 8
For Y = (minY + 8) To maxY - 8
    ScreenX = 8
    For X = minX + 8 To maxX - 8
        If X > 100 Or Y < 1 Then Exit For
        'Layer 1 **********************************
        With MapData(X, Y).Graphic(1)
            If (.Started = 1) Then
                .FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed)
                If .FrameCounter > GrhData(.GrhIndex).NumFrames Then
                    .FrameCounter = (.FrameCounter Mod GrhData(.GrhIndex).NumFrames) + 1
                End If
            End If

            'Figure out what frame to draw (always 1 if not animated)
            iGrhIndex = GrhData(.GrhIndex).Frames(.FrameCounter)
        End With

        rSourceRect.Left = GrhData(iGrhIndex).sX
        rSourceRect.Top = GrhData(iGrhIndex).sY
        rSourceRect.Right = rSourceRect.Left + GrhData(iGrhIndex).pixelWidth
        rSourceRect.Bottom = rSourceRect.Top + GrhData(iGrhIndex).pixelHeight

        'El width fue hardcodeado para speed!
        Call BackBufferSurface.BltFast( _
                ((32 * ScreenX) - 32) + PixelOffsetX, _
                ((32 * ScreenY) - 32) + PixelOffsetY, _
                SurfaceDB.Surface(GrhData(iGrhIndex).FileNum), _
                rSourceRect, _
                DDBLTFAST_WAIT)
        '******************************************
        'Layer 2 **********************************
        If MapData(X, Y).Graphic(2).GrhIndex <> 0 Then
            Call DDrawTransGrhtoSurface( _
                    BackBufferSurface, _
                    MapData(X, Y).Graphic(2), _
                    ((32 * ScreenX) - 32) + PixelOffsetX, _
                    ((32 * ScreenY) - 32) + PixelOffsetY, _
                    1, _
                    1)
        End If
        '******************************************
        ScreenX = ScreenX + 1
    Next X
    ScreenY = ScreenY + 1
    If Y > 100 Then Exit For
Next Y


'busco que nombre dibujar
Call ConvertCPtoTP(frmMain.MainViewShp.Left, frmMain.MainViewShp.Top, frmMain.MouseX, frmMain.MouseY, nX, nY)


'Draw Transparent Layers  (Layer 2, 3)
ScreenY = 8
For Y = minY + 8 To maxY - 1
    ScreenX = 5
    For X = minX + 5 To maxX - 5
        If X > 100 Or X < -3 Then Exit For
        iPPx = 32 * ScreenX - 32 + PixelOffsetX
        iPPy = 32 * ScreenY - 32 + PixelOffsetY

        'Object Layer **********************************
        If Abs(nX - X) < 1 And (Abs(nY - Y)) < 1 Then
            If MapData(X, Y).ObjGrh.GrhIndex <> 0 Then
                    Call SurfaceColor( _
                            BackBufferSurface, _
                            MapData(X, Y).ObjGrh, _
                            iPPx, iPPy, 1, 1)
            End If
        Else
            If MapData(X, Y).ObjGrh.GrhIndex <> 0 Then
                    Call DDrawTransGrhtoSurface( _
                            BackBufferSurface, _
                            MapData(X, Y).ObjGrh, _
                            iPPx, iPPy, 1, 1)
            End If
        End If
        '***********************************************
        'Char layer ************************************
        If MapData(X, Y).CharIndex <> 0 Then
            tempChar = charlist(MapData(X, Y).CharIndex)
            PixelOffsetXTemp = PixelOffsetX
            PixelOffsetYTemp = PixelOffsetY


            Moved = 0
            With tempChar
                If .scrollDirectionX <> 0 Then
                    .MoveOffsetX = .MoveOffsetX + 8 * Sgn(.scrollDirectionX) * timerTicksPerFrame

                    .Body.Walk(.Heading).Started = 1
                    .Arma.WeaponWalk(.Heading).Started = 1
                    .Escudo.ShieldWalk(.Heading).Started = 1

                    'Char moved
                    Moved = 1
                        
                    'Check if we already got there
                    If (Sgn(.scrollDirectionX) = 1 And .MoveOffsetX >= 0) Or _
                      (Sgn(.scrollDirectionX) = -1 And .MoveOffsetX <= 0) Then
                        .MoveOffsetX = 0
                        .scrollDirectionX = 0
                    End If
                    
                    PixelOffsetXTemp = PixelOffsetXTemp + .MoveOffsetX
                End If
                
                'If needed, move up and down
                If .scrollDirectionY <> 0 Then
                    .MoveOffsetY = .MoveOffsetY + 8 * Sgn(.scrollDirectionY) * timerTicksPerFrame
                        
                    .Body.Walk(.Heading).Started = 1
                    .Arma.WeaponWalk(.Heading).Started = 1
                    .Escudo.ShieldWalk(.Heading).Started = 1

                    Moved = 1
                        
                    'Check if we already got there
                    If (Sgn(.scrollDirectionY) = 1 And .MoveOffsetY >= 0) Or _
                      (Sgn(.scrollDirectionY) = -1 And .MoveOffsetY <= 0) Then
                        .MoveOffsetY = 0
                        .scrollDirectionY = 0
                    End If
                    PixelOffsetYTemp = PixelOffsetYTemp + .MoveOffsetY
                End If
            End With
            'If done moving stop animation
            If Moved = 0 And tempChar.Moving = 1 Then
                tempChar.Moving = 0
                tempChar.Body.Walk(tempChar.Heading).FrameCounter = 1
                tempChar.Body.Walk(tempChar.Heading).Started = 0
                tempChar.Arma.WeaponWalk(tempChar.Heading).FrameCounter = 1
                tempChar.Arma.WeaponWalk(tempChar.Heading).Started = 0
                tempChar.Escudo.ShieldWalk(tempChar.Heading).FrameCounter = 1
                tempChar.Escudo.ShieldWalk(tempChar.Heading).Started = 0
            End If
            
            '[ANIM ATAK]
            If tempChar.Arma.WeaponAttack > 0 Then
                tempChar.Arma.WeaponAttack = tempChar.Arma.WeaponAttack - 1
                If tempChar.Arma.WeaponAttack = 0 Then
                    tempChar.Arma.WeaponWalk(tempChar.Heading).Started = 0
                End If
            End If
            '[/ANIM ATAK]
            
            'Dibuja solamente players
            iPPx = ((32 * ScreenX) - 32) + PixelOffsetXTemp
            iPPy = ((32 * ScreenY) - 32) + PixelOffsetYTemp
            If tempChar.Head.Head(tempChar.Heading).GrhIndex <> 0 Or (UCase$(tempChar.Nombre) = UCase$(UserName) Or mid(tempChar.Nombre, InStr(tempChar.Nombre, "<")) And UserNavegando = True) Then
                If Not charlist(MapData(X, Y).CharIndex).invisible Then
                #If (ConAlfaB = 1) Then
                        If tempChar.Aura.GrhIndex Then
                            Call DDrawTransGrhtoSurfaceAlpha(BackBufferSurface, tempChar.Aura, _
                                    (((32 * ScreenX) - 32) + PixelOffsetXTemp), _
                                    (((35 * ScreenY) - 35) + PixelOffsetYTemp), _
                                    1, 0, 1)
    #End If
End If

                        '[CUERPO]'
                            Call DDrawTransGrhtoSurface(BackBufferSurface, tempChar.Body.Walk(tempChar.Heading), _
                                    (((32 * ScreenX) - 32) + PixelOffsetXTemp), _
                                    (((32 * ScreenY) - 32) + PixelOffsetYTemp), _
                                    1, 1)
                        '[CABEZA]'
                            Call DDrawTransGrhtoSurface( _
                                    BackBufferSurface, _
                                    tempChar.Head.Head(tempChar.Heading), _
                                    iPPx + tempChar.Body.HeadOffset.X, _
                                    iPPy + tempChar.Body.HeadOffset.Y, _
                                    1, 0)
                        '[Casco]'
                            If tempChar.Casco.Head(tempChar.Heading).GrhIndex <> 0 Then
                                Call DDrawTransGrhtoSurface( _
                                        BackBufferSurface, _
                                        tempChar.Casco.Head(tempChar.Heading), _
                                        iPPx + tempChar.Body.HeadOffset.X, _
                                        iPPy + tempChar.Body.HeadOffset.Y, _
                                        1, 0)
                            End If
                        '[ARMA]'
                            If tempChar.Arma.WeaponWalk(tempChar.Heading).GrhIndex <> 0 Then
                                Call DDrawTransGrhtoSurface( _
                                        BackBufferSurface, _
                                        tempChar.Arma.WeaponWalk(tempChar.Heading), _
                                        iPPx, iPPy, 1, 1)
                            End If
                        '[Escudo]'
                            If tempChar.Escudo.ShieldWalk(tempChar.Heading).GrhIndex <> 0 Then
                                Call DDrawTransGrhtoSurface( _
                                        BackBufferSurface, _
                                        tempChar.Escudo.ShieldWalk(tempChar.Heading), _
                                        iPPx, iPPy, 1, 1)
                            End If
                    
                    
 If Nombres Then
                    If tempChar.invisible = False Then
                        If tempChar.Nombre <> "" Then
                            Dim lCenter As Long
                            If InStr(tempChar.Nombre, "<") > 0 And InStr(tempChar.Nombre, ">") > 0 Then
                                lCenter = (frmMain.TextWidth(Left(tempChar.Nombre, InStr(tempChar.Nombre, "<") - 1)) / 2) - 16
                                Dim sClan As String: sClan = mid(tempChar.Nombre, InStr(tempChar.Nombre, "<"))
                                Dim ColorClan As Long
                                ColorClan = RGB(231, 202, 157)
                               
                                Select Case tempChar.priv
                                Case 0
                                    If tempChar.Criminal Then
                                        Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, Left(tempChar.Nombre, InStr(tempChar.Nombre, "<") - 1), RGB(ColoresPJ(50).r, ColoresPJ(50).g, ColoresPJ(50).b))
                                        lCenter = (frmMain.TextWidth(sClan) / 2) - 16
                                        Call Dialogos.DrawText(iPPx - lCenter, iPPy + 45, sClan, ColorClan)
                                    Else
                                        Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, Left(tempChar.Nombre, InStr(tempChar.Nombre, "<") - 1), RGB(ColoresPJ(49).r, ColoresPJ(49).g, ColoresPJ(49).b))
                                        lCenter = (frmMain.TextWidth(sClan) / 2) - 16
                                        Call Dialogos.DrawText(iPPx - lCenter, iPPy + 45, sClan, ColorClan)
                                    End If
                                Case 25  'admin
                                    Call Dialogos.DrawTextBig(iPPx - lCenter, iPPy + 30, Left(tempChar.Nombre, InStr(tempChar.Nombre, "<") - 1), RGB(ColoresPJ(tempChar.priv).r, ColoresPJ(tempChar.priv).g, ColoresPJ(tempChar.priv).b))
                                    lCenter = (frmMain.TextWidth(sClan) / 2) - 16
                                    Call Dialogos.DrawTextBig(iPPx - lCenter, iPPy + 45, sClan, ColorClan)
                                Case Else 'el resto
                                    Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, Left(tempChar.Nombre, InStr(tempChar.Nombre, "<") - 1), RGB(ColoresPJ(tempChar.priv).r, ColoresPJ(tempChar.priv).g, ColoresPJ(tempChar.priv).b))
                                    lCenter = (frmMain.TextWidth(sClan) / 2) - 16
                                    Call Dialogos.DrawText(iPPx - lCenter, iPPy + 45, sClan, ColorClan)
                                End Select
                            Else
                                lCenter = (frmMain.TextWidth(tempChar.Nombre) / 2) - 16
 
                                Select Case tempChar.priv
                                Case 0
                                    If tempChar.Criminal Then
                                        Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, tempChar.Nombre, RGB(ColoresPJ(50).r, ColoresPJ(50).g, ColoresPJ(50).b))
                                    Else
                                        Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, tempChar.Nombre, RGB(ColoresPJ(49).r, ColoresPJ(49).g, ColoresPJ(49).b))
                                    End If
                                Case 7
                                    Call Dialogos.DrawTextBig(iPPx - lCenter, iPPy + 30, tempChar.Nombre, RGB(ColoresPJ(tempChar.priv).r, ColoresPJ(tempChar.priv).g, ColoresPJ(tempChar.priv).b))
                                Case Else
                                    Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, tempChar.Nombre, RGB(ColoresPJ(tempChar.priv).r, ColoresPJ(tempChar.priv).g, ColoresPJ(tempChar.priv).b))
                                End Select
                            End If
                        End If
                        End If
                        End If

                End If  'end if ~in

                If Dialogos.CantidadDialogos > 0 Then
                    Call Dialogos.Update_Dialog_Pos( _
                            (iPPx + tempChar.Body.HeadOffset.X), _
                            (iPPy + tempChar.Body.HeadOffset.Y), _
                            MapData(X, Y).CharIndex)
                End If
                
                
            Else '<-> If TempChar.Head.Head(TempChar.Heading).GrhIndex <> 0 Then
                If Dialogos.CantidadDialogos > 0 Then
                    Call Dialogos.Update_Dialog_Pos( _
                            (iPPx + tempChar.Body.HeadOffset.X), _
                            (iPPy + tempChar.Body.HeadOffset.Y), _
                            MapData(X, Y).CharIndex)
                End If

                Call DDrawTransGrhtoSurface( _
                        BackBufferSurface, _
                        tempChar.Body.Walk(tempChar.Heading), _
                        iPPx, iPPy, 1, 1)
            End If '<-> If TempChar.Head.Head(TempChar.Heading).GrhIndex <> 0 Then


            'Refresh charlist
            charlist(MapData(X, Y).CharIndex) = tempChar

            'BlitFX (TM)
            If charlist(MapData(X, Y).CharIndex).Fx <> 0 Then
#If (ConAlfaB = 1) Then
                Call DDrawTransGrhtoSurfaceAlpha( _
                        BackBufferSurface, _
                        FxData(tempChar.Fx).Fx, _
                        iPPx + FxData(tempChar.Fx).OffsetX, _
                        iPPy + FxData(tempChar.Fx).OffsetY, _
                        1, 1, MapData(X, Y).CharIndex)
#Else
                Call DDrawTransGrhtoSurface( _
                        BackBufferSurface, _
                        FxData(tempChar.Fx).Fx, _
                        iPPx + FxData(tempChar.Fx).OffsetX, _
                        iPPy + FxData(tempChar.Fx).OffsetY, _
                        1, 1, MapData(X, Y).CharIndex)
#End If
            End If
        End If '<-> If MapData(X, Y).CharIndex <> 0 Then
        '*************************************************
        'Layer 3 *****************************************
        If MapData(X, Y).Graphic(3).GrhIndex <> 0 Then
            'Draw
            Call DDrawTransGrhtoSurface( _
                    BackBufferSurface, _
                    MapData(X, Y).Graphic(3), _
                    ((32 * ScreenX) - 32) + PixelOffsetX, _
                    ((32 * ScreenY) - 32) + PixelOffsetY, _
                    1, 1)
        End If
        If Abs(nX - X) < 1 And (Abs(nY - Y)) < 1 Then
    If MapData(X, Y).ObjGrh.GrhIndex <> 0 Then
        Dialogos.DrawText frmMain.MouseX + 240, frmMain.MouseY + 100, MapData(X, Y).ObjName, vbWhite
    End If
End If
        '************************************************
        ScreenX = ScreenX + 1
    Next X
    ScreenY = ScreenY + 1
    If Y >= 100 Or Y < 1 Then Exit For
Next Y

If Not bTecho Then
    'Draw blocked tiles and grid
    ScreenY = 5
    For Y = minY + 5 To maxY - 1
        ScreenX = 5
        For X = minX + 5 To maxX
            'Check to see if in bounds
            If X < 101 And X > 0 And Y < 101 And Y > 0 Then
                If MapData(X, Y).Graphic(4).GrhIndex <> 0 Then
                    'Draw
                    Call DDrawTransGrhtoSurface( _
                        BackBufferSurface, _
                        MapData(X, Y).Graphic(4), _
                        ((32 * ScreenX) - 32) + PixelOffsetX, _
                        ((32 * ScreenY) - 32) + PixelOffsetY, _
                        1, 1)
                End If
            End If
            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y
End If



Dim PP As RECT

PP.Left = 0
PP.Top = 0
PP.Right = WindowTileWidth * TilePixelWidth
PP.Bottom = WindowTileHeight * TilePixelHeight

'*************************
'*****Lorwik - Noche******
'*************************

#If ConAlfaB Then
'Efectos
If EfectosDiaY Then
If Anochecer = 1 Then
EfectoNoche BackBufferSurface
End If
If Atardecer = 1 Then
EfectoTarde BackBufferSurface
End If
If Amanecer = 1 Then
EfectoAmanecer BackBufferSurface
End If
End If
#End If

'Lorwik - Macros
Call CargarMacros

            frmMain.Coord.Caption = UserMap
            frmMain.Coord2.Caption = UserPos.X
            frmMain.coord3.Caption = UserPos.Y

End Sub
Public Function RenderSounds()
'Lorwik> todo para una fogata de mierda -.-
    DoFogataFx
End Function


Function HayUserAbajo(ByVal X As Integer, ByVal Y As Integer, ByVal GrhIndex As Integer) As Boolean

If GrhIndex > 0 Then
        
        HayUserAbajo = _
            charlist(UserCharIndex).Pos.X >= X - (GrhData(GrhIndex).TileWidth \ 2) _
        And charlist(UserCharIndex).Pos.X <= X + (GrhData(GrhIndex).TileWidth \ 2) _
        And charlist(UserCharIndex).Pos.Y >= Y - (GrhData(GrhIndex).TileHeight - 1) _
        And charlist(UserCharIndex).Pos.Y <= Y
        
End If
End Function

Function PixelPos(ByVal X As Integer) As Integer
'*****************************************************************
'Converts a tile position to a screen position
'*****************************************************************
    PixelPos = (TilePixelWidth * X) - TilePixelWidth
End Function

Sub LoadGraphics()
    Call SurfaceDB.Initialize(DirectDraw, ClientSetup.bUseVideo, DirGraficos, ClientSetup.byMemory)
     Call frmCargando.progresoConDelay(75)
End Sub

'[END]'
Function InitTileEngine(ByRef setDisplayFormhWnd As Long, setMainViewTop As Integer, setMainViewLeft As Integer, setTilePixelHeight As Integer, setTilePixelWidth As Integer, setWindowTileHeight As Integer, setWindowTileWidth As Integer, setTileBufferSize As Integer) As Boolean
'*****************************************************************
'InitEngine
'*****************************************************************
Dim SurfaceDesc As DDSURFACEDESC2
Dim ddck As DDCOLORKEY

IniPath = App.Path & "\Init\"

'Set intial user position
UserPos.X = MinXBorder
UserPos.Y = MinYBorder

'Fill startup variables

DisplayFormhWnd = setDisplayFormhWnd
MainViewTop = setMainViewTop
MainViewLeft = setMainViewLeft
TilePixelWidth = setTilePixelWidth
TilePixelHeight = setTilePixelHeight
WindowTileHeight = setWindowTileHeight
WindowTileWidth = setWindowTileWidth
TileBufferSize = setTileBufferSize

MinXBorder = XMinMapSize + (WindowTileWidth \ 2)
MaxXBorder = XMaxMapSize - (WindowTileWidth \ 2)
MinYBorder = YMinMapSize + (WindowTileHeight \ 2)
MaxYBorder = YMaxMapSize - (WindowTileHeight \ 2)

MainViewWidth = (TilePixelWidth * WindowTileWidth)
MainViewHeight = (TilePixelHeight * WindowTileHeight)

'Lorwik> Antes estaba en el sub main
    MainViewRect.Left = MainViewLeft
    MainViewRect.Top = MainViewTop
    MainViewRect.Right = MainViewRect.Left + MainViewWidth
    MainViewRect.Bottom = MainViewRect.Top + MainViewHeight
    
    MainDestRect.Left = TilePixelWidth * TileBufferSize - TilePixelWidth
    MainDestRect.Top = TilePixelHeight * TileBufferSize - TilePixelHeight
    MainDestRect.Right = MainDestRect.Left + MainViewWidth
    MainDestRect.Bottom = MainDestRect.Top + MainViewHeight

ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock





DirectDraw.SetCooperativeLevel DisplayFormhWnd, DDSCL_NORMAL

'Primary Surface
' Fill the surface description structure
With SurfaceDesc
    .lFlags = DDSD_CAPS
    .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
End With



Set PrimarySurface = DirectDraw.CreateSurface(SurfaceDesc)

Set PrimaryClipper = DirectDraw.CreateClipper(0)
PrimaryClipper.SetHWnd frmMain.hWnd
PrimarySurface.SetClipper PrimaryClipper

Set SecundaryClipper = DirectDraw.CreateClipper(0)

With BackBufferRect
    .Left = 0
    .Top = 0
    .Right = TilePixelWidth * (WindowTileWidth + 2 * TileBufferSize)
    .Bottom = TilePixelHeight * (WindowTileHeight + 2 * TileBufferSize)
End With

With SurfaceDesc
    .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    If ClientSetup.bUseVideo Then
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    Else
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    End If
    .lHeight = BackBufferRect.Bottom
    .lWidth = BackBufferRect.Right
End With

Set BackBufferSurface = DirectDraw.CreateSurface(SurfaceDesc)

ddck.Low = 0
ddck.High = 0
BackBufferSurface.SetColorKey DDCKEY_SRCBLT, ddck



Call LoadGrhData
Call CargarCuerpos
Call CargarCabezas
Call CargarCascos
Call CargarFxs
Call frmCargando.progresoConDelay(40)

Call LoadGraphics

InitTileEngine = True

End Function

Sub ShowNextFrame()
'***********************************************
'Updates and draws next frame to screen
'***********************************************
    Static OffsetCounterX As Single
    Static OffsetCounterY As Single
    
    '****** Set main view rectangle ******
    GetWindowRect DisplayFormhWnd, MainViewRect
    
    With MainViewRect
        .Left = .Left + MainViewLeft
        .Top = .Top + MainViewTop
        .Right = .Left + MainViewWidth
        .Bottom = .Top + MainViewHeight
    End With
    
    If EngineRun Then
        '****** Move screen Left and Right if needed ******
        If AddtoUserPos.X <> 0 Then
            OffsetCounterX = OffsetCounterX - 8 * AddtoUserPos.X * timerTicksPerFrame
            If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.X) Then
                OffsetCounterX = 0
                AddtoUserPos.X = 0
                UserMoving = 0
            End If
        '****** Move screen Up and Down if needed ******
        ElseIf AddtoUserPos.Y <> 0 Then
            OffsetCounterY = OffsetCounterY - 8 * AddtoUserPos.Y * timerTicksPerFrame
            If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.Y) Then
                OffsetCounterY = 0
                AddtoUserPos.Y = 0
                UserMoving = 0
            End If
        End If

        '****** Update screen ******
        Call RenderScreen(UserPos.X - AddtoUserPos.X, UserPos.Y - AddtoUserPos.Y, OffsetCounterX, OffsetCounterY)

        If IScombate Then
        frmMain.combate.Visible = True
        frmMain.combateII.Visible = False
        Else
        frmMain.combate.Visible = False
        frmMain.combateII.Visible = True
        End If
        
        'Estado del Dia
#If ConAlfaB Then
        If Amanecer Then
            frmMain.Dia.ForeColor = RGB(255, 128, 64)
            frmMain.Dia.Caption = "Amanecer"
            frmMain.Clima = General_Load_Picture_From_Resource("[Main]ClimaMañana.gif")
        ElseIf Atardecer Then
            frmMain.Dia.ForeColor = RGB(250, 61, 5)
            frmMain.Dia.Caption = "Tarde"
            frmMain.Clima = General_Load_Picture_From_Resource("[Main]ClimaTarde.gif")
        ElseIf Anochecer Then
            frmMain.Dia.ForeColor = RGB(128, 128, 128)
            frmMain.Dia.Caption = "Noche"
            frmMain.Clima = General_Load_Picture_From_Resource("[Main]ClimaNoche.gif")
        Else
            frmMain.Dia.ForeColor = RGB(0, 255, 0)
            frmMain.Dia.Caption = "Dia"
            frmMain.Clima = General_Load_Picture_From_Resource("[Main]ClimaDia.gif")
        End If
#End If
        
        If CartelParalisis Then Call Dialogos.DrawText(260, 310, CartelParalisis & " segundos restantes de paralisis.", vbCyan)
                If CartelInvisibilidad Then Call Dialogos.DrawText(260, 300, CartelInvisibilidad & " segundos restantes de Invisibilidad", vbCyan)
                If UserGuerra Then Call Dialogos.DrawText(260, 270, "¡Estas En Guerra!", vbRed) 'Guerras
        Call Dialogos.MostrarTexto
        Call DibujarCartel
        Call Dialogos.DrawUsers
        Call DialogosClanes.Draw(Dialogos)
        
        Call DrawBackBufferSurface
        
        FramesPerSecCounter = FramesPerSecCounter + 1
    End If
End Sub

Sub CrearGrh(GrhIndex As Integer, Index As Integer)
ReDim Preserve Grh(1 To Index) As Grh
Grh(Index).FrameCounter = 1
Grh(Index).GrhIndex = GrhIndex
Grh(Index).Speed = GrhData(GrhIndex).Speed
Grh(Index).Started = 1
End Sub

Sub CargarAnimsExtra()
Call CrearGrh(6580, 1) 'Anim Invent
Call CrearGrh(534, 2) 'Animacion de teleport
End Sub

Function ControlVelocidad(ByVal lastTime As Long) As Boolean
ControlVelocidad = (GetTickCount - lastTime > 20)
End Function


#If ConAlfaB Then



#End If
#If ConAlfaB Then
Sub SurfaceColor(Surface As DirectDrawSurface7, Grh As Grh, ByVal X As Integer, ByVal Y As Integer, center As Byte, Animate As Byte, Optional ByVal KillAnim As Integer = 0)
 
Dim iGrhIndex As Integer
Dim SourceRect As RECT
Dim QuitarAnimacion As Boolean
 
 
If Animate Then
    If Grh.Started = 1 Then
        Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed)
        If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
            Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
            If KillAnim Then
                If charlist(KillAnim).FxLoopTimes <> LoopAdEternum Then
                    If charlist(KillAnim).FxLoopTimes > 0 Then charlist(KillAnim).FxLoopTimes = charlist(KillAnim).FxLoopTimes - 1
                    If charlist(KillAnim).FxLoopTimes < 1 Then 'Matamos la anim del fx ;))
                        charlist(KillAnim).Fx = 0
                        Exit Sub
                    End If
                End If
            End If
        End If
    End If
End If
 
If Grh.GrhIndex = 0 Then Exit Sub
 
iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
 
If center Then
    If GrhData(iGrhIndex).TileWidth <> 1 Then
        X = X - Int(GrhData(iGrhIndex).TileWidth * 16) + 16 'hard coded for speed
    End If
    If GrhData(iGrhIndex).TileHeight <> 1 Then
        Y = Y - Int(GrhData(iGrhIndex).TileHeight * 32) + 32 'hard coded for speed
    End If
End If
 
With SourceRect
    .Left = GrhData(iGrhIndex).sX + IIf(X < 0, Abs(X), 0)
    .Top = GrhData(iGrhIndex).sY + IIf(Y < 0, Abs(Y), 0)
    .Right = .Left + GrhData(iGrhIndex).pixelWidth
    .Bottom = .Top + GrhData(iGrhIndex).pixelHeight
End With
 
Dim src As DirectDrawSurface7
Dim rDest As RECT
Dim dArray() As Byte, sArray() As Byte
Dim ddsdSrc As DDSURFACEDESC2, ddsdDest As DDSURFACEDESC2
Dim Modo As Long
 
Set src = SurfaceDB.Surface(GrhData(iGrhIndex).FileNum)
 
src.GetSurfaceDesc ddsdSrc
Surface.GetSurfaceDesc ddsdDest
With rDest
    .Left = X
    .Top = Y
    .Right = X + GrhData(iGrhIndex).pixelWidth
    .Bottom = Y + GrhData(iGrhIndex).pixelHeight
   
    If .Right > ddsdDest.lWidth Then
        .Right = ddsdDest.lWidth
    End If
    If .Bottom > ddsdDest.lHeight Then
        .Bottom = ddsdDest.lHeight
    End If
End With
 
Dim SrcLock As Boolean, DstLock As Boolean
SrcLock = False: DstLock = False
 
On Local Error GoTo HayErrorAlpha
 
src.Lock SourceRect, ddsdSrc, DDLOCK_NOSYSLOCK Or DDLOCK_WAIT, 0
Surface.Lock rDest, ddsdDest, DDLOCK_NOSYSLOCK Or DDLOCK_WAIT, 0
 
Surface.GetLockedArray dArray()
src.GetLockedArray sArray()
       
If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0 Then
  Modo = 555
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 Then
  Modo = 565
Else
  MsgBox "Modo de vídeo no esta en 555 o 565 o algo falló."
  End
End If
 
Call vbDABLcolorblend16565ck(ByVal VarPtr(sArray(SourceRect.Left * 2, SourceRect.Top)), ByVal VarPtr(dArray(X + X, Y)), 150, rDest.Right - rDest.Left, rDest.Bottom - rDest.Top, ddsdSrc.lPitch, ddsdDest.lPitch, 255, 230, 138)
Surface.Unlock rDest
src.Unlock SourceRect
 
Exit Sub
 
HayErrorAlpha:
If SrcLock Then src.Unlock SourceRect
If DstLock Then Surface.Unlock rDest
 
End Sub
#End If
#If ConAlfaB Then
 
Public Sub EfectoNoche(ByRef Surface As DirectDrawSurface7)
Dim dArray() As Byte, sArray() As Byte
Dim ddsdDest As DDSURFACEDESC2
Dim Modo As Long
Dim rRect As RECT
 
Surface.GetSurfaceDesc ddsdDest
 
With rRect
.Left = 0
.Top = 0
.Right = ddsdDest.lWidth
.Bottom = ddsdDest.lHeight
End With
 
If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0 Then
Modo = 0
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 Then
Modo = 1
Else
Modo = 2
End If
 
Dim DstLock As Boolean
DstLock = False
 
On Local Error GoTo HayErrorAlpha
 
Surface.Lock rRect, ddsdDest, DDLOCK_WAIT, 0
DstLock = True
 
Surface.GetLockedArray dArray()
Call BltEfectoNoche(ByVal VarPtr(dArray(0, 0)), _
ddsdDest.lWidth, ddsdDest.lHeight, ddsdDest.lPitch, _
Modo)
 
HayErrorAlpha:
 
If DstLock = True Then
Surface.Unlock rRect
DstLock = False
End If
 
End Sub
 
Public Sub EfectoTarde(ByRef Surface As DirectDrawSurface7)
Dim dArray() As Byte, sArray() As Byte
Dim ddsdDest As DDSURFACEDESC2
Dim Modo As Long
Dim rRect As RECT
 
Surface.GetSurfaceDesc ddsdDest
 
With rRect
.Left = 0
.Top = 0
.Right = ddsdDest.lWidth
.Bottom = ddsdDest.lHeight
End With
 
If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0 Then
Modo = 0
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 Then
Modo = 1
Else
Modo = 2
End If
 
Dim DstLock As Boolean
DstLock = False
 
On Local Error GoTo HayErrorAlpha
 
Surface.Lock rRect, ddsdDest, DDLOCK_WAIT, 0
DstLock = True
 
Surface.GetLockedArray dArray()
 
 
Call vbDABLcolorblend16565ck(ByVal VarPtr(dArray(0, 0)), ByVal VarPtr(dArray(0, 0)), 60, rRect.Right - rRect.Left, rRect.Bottom - rRect.Top, ddsdDest.lPitch, ddsdDest.lPitch, 0, 0, 0)
 
HayErrorAlpha:
 
If DstLock = True Then
Surface.Unlock rRect
DstLock = False
End If
 
End Sub
 
Public Sub EfectoAmanecer(ByRef Surface As DirectDrawSurface7)
Dim dArray() As Byte, sArray() As Byte
Dim ddsdDest As DDSURFACEDESC2
Dim Modo As Long
Dim rRect As RECT
 
Surface.GetSurfaceDesc ddsdDest
 
With rRect
.Left = 0
.Top = 0
.Right = ddsdDest.lWidth
.Bottom = ddsdDest.lHeight
End With
 
If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0 Then
Modo = 0
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 Then
Modo = 1
Else
Modo = 2
End If
 
Dim DstLock As Boolean
DstLock = False
 
On Local Error GoTo HayErrorAlpha
 
Surface.Lock rRect, ddsdDest, DDLOCK_WAIT, 0
DstLock = True
 
Surface.GetLockedArray dArray()
 
 
Call vbDABLcolorblend16565ck(ByVal VarPtr(dArray(0, 0)), ByVal VarPtr(dArray(0, 0)), 70, rRect.Right - rRect.Left, rRect.Bottom - rRect.Top, ddsdDest.lPitch, ddsdDest.lPitch, 128, 64, 64)
HayErrorAlpha:
 
If DstLock = True Then
Surface.Unlock rRect
DstLock = False
End If
 
End Sub
#End If

