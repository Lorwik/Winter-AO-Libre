Attribute VB_Name = "Mod_TileEngine"
Option Explicit

Dim map_current As Map
Dim char_list() As Char

'Screen positioning
Public minY As Integer          'Start Y pos on current screen + tilebuffer
Public maxY As Integer          'End Y pos on current screen
Public minX As Integer          'Start X pos on current screen
Public maxX As Integer          'End X pos on current screen

Public movSpeed As Single

'Map sizes in tiles
Public Const XMaxMapSize As Byte = 100
Public Const XMinMapSize As Byte = 1
Public Const YMaxMapSize As Byte = 100
Public Const YMinMapSize As Byte = 1

''
'Sets a Grh animation to loop indefinitely.
Private Const INFINITE_LOOPS As Integer = -1

'Encabezado bmp
Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type

'Info del encabezado del bmp
Private Type BITMAPINFOHEADER
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
    
    Active As Boolean
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

Private Type Particle
    friction As Single
    X As Single
    Y As Single
    vector_x As Single
    vector_y As Single
    angle As Single
    Grh As Grh
    alive_counter As Long
    x1 As Long
    x2 As Long
    y1 As Long
    y2 As Long
    vecx1 As Long
    vecx2 As Long
    vecy1 As Long
    vecy2 As Long
    life1 As Long
    life2 As Long
    fric As Long
    spin_speedL As Single
    spin_speedH As Single
    gravity As Boolean
    grav_strength As Long
    bounce_strength As Long
    spin As Boolean
    XMove As Boolean
    YMove As Boolean
    move_x1 As Integer
    move_x2 As Integer
    move_y1 As Integer
    move_y2 As Integer
    Radio As Integer
    rgb_list(0 To 3) As Long
End Type
 
Dim base_tile_size As Integer

Private Type decoration
    Grh As Grh
    Render_On_Top As Boolean
    subtile_pos As Byte
End Type

Private Type Map_Tile
    Grh(1 To 3) As Grh
    decoration(1 To 5) As decoration
    decoration_count As Byte
    blocked As Boolean
    particle_group_index As Long
    char_index As Long
    light_base_value(0 To 3) As Long
    light_value(0 To 3) As Long
   
    exit_index As Long
    npc_index As Long
    item_index As Long
   
    Trigger As Byte
End Type

Private Type Map
    map_grid() As Map_Tile
    map_x_max As Long
    map_x_min As Long
    map_y_max As Long
    map_y_min As Long
    map_description As String
    'Added by Juan Martín Sotuyo Dodero
    base_light_color As Long
End Type

'*********************************
'Particulas
'*********************************
Private Type Stream
    name As String
    NumOfParticles As Long
    NumGrhs As Long
    ID As Long
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
   
    Speed As Single
    life_counter As Long
End Type

Private Type particle_group
    Active As Boolean
    ID As Long
    map_x As Long
    map_y As Long
    char_index As Long

    frame_counter As Single
    frame_speed As Single
    
    stream_type As Byte

    particle_stream() As Particle
    particle_count As Long
    
    grh_index_list() As Long
    grh_index_count As Long
    
    alpha_blend As Boolean
    
    alive_counter As Long
    never_die As Boolean
    
    live As Long
    liv1 As Integer
    liveend As Long
    
    x1 As Long
    x2 As Long
    y1 As Long
    y2 As Long
    angle As Long
    vecx1 As Long
    vecx2 As Long
    vecy1 As Long
    vecy2 As Long
    life1 As Long
    life2 As Long
    fric As Long
    spin_speedL As Single
    spin_speedH As Single
    gravity As Boolean
    grav_strength As Long
    bounce_strength As Long
    spin As Boolean
    XMove As Boolean
    YMove As Boolean
    move_x1 As Long
    move_x2 As Long
    move_y1 As Long
    move_y2 As Long
    rgb_list(0 To 3) As Long
    
    Speed As Single
    life_counter As Long
    
    Radio As Integer
End Type
'Particle system
 
'Dim StreamData() As particle_group
Dim TotalStreams As Long
Dim particle_group_list() As particle_group
Dim particle_group_count As Long
Dim particle_group_last As Long

'*********************************
'*********************************

'Lista de cuerpos
Public Type BodyData
    Walk(E_Heading.NORTH To E_Heading.WEST) As Grh
    HeadOffset As Position
End Type

'Lista de cabezas
Public Type HeadData
    Head(E_Heading.NORTH To E_Heading.WEST) As Grh
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

'Apariencia del personaje
Public Type Char
    Active As Byte
    Heading As E_Heading
    Pos As Position
    
    iHead As Integer
    iBody As Integer
    Body As BodyData
    Head As HeadData
    casco As HeadData
    arma As WeaponAnimData
    escudo As ShieldAnimData
    UsandoArma As Boolean
    
    fX As Grh
    FxIndex As Integer
    
    Aura_Index As Integer
    Aura As Grh
    
    ParticulaIndex As Integer 'se usa

    particle_count As Integer
    particle_group() As Long
    
    Criminal As Byte
    Atacable As Boolean
    
    nombre As String
    
    scrollDirectionX As Integer
    scrollDirectionY As Integer
    
    Moving As Byte
    MoveOffsetX As Single
    MoveOffsetY As Single
    
    pie As Boolean
    Muerto As Boolean
    invisible As Boolean
    priv As Byte
End Type

'Info de un objeto
Public Type Obj
    OBJIndex As Integer
    Amount As Integer
End Type

'Sistema de sangre
Public Type BloodPool
    LifeTime As Integer
    Active As Byte
    Grh As Grh
    Alpha As Integer
    color(3) As Long
    Head As E_Heading
End Type

'Tipo de las celdas del mapa
Public Type MapBlock
    particle_group_index As Long
    Graphic(1 To 4) As Grh
    CharIndex As Integer
    ObjGrh As Grh
    
    NPCIndex As Integer
    OBJInfo As Obj
    TileExit As WorldPos
    blocked As Byte
    
    light_value(3) As Long
    base_light(0 To 3) As Boolean 'Indica si el tile tiene luz propia.
    
    color(3) As Long
    
    Trigger As Integer
    Blood As BloodPool
End Type

'Info de cada mapa
Public Type MapInfo
    Music As String
    name As String
    StartPos As WorldPos
    MapVersion As Integer
End Type

'DX8 Objects
Public DirectX As New DirectX8
Public DirectD3D8 As D3DX8
Public DirectD3D As Direct3D8
Public DirectDevice As Direct3DDevice8

Public Type TLVERTEX
    X As Single
    Y As Single
    z As Single
    rhw As Single
    color As Long
    Specular As Long
    tu As Single
    tv As Single
End Type

'Bordes del mapa
Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

Public UserIndex As Integer
Public UserMoving As Byte
Public UserBody As Integer
Public UserHead As Integer
Public UserPos As Position 'Posicion
Public AddtoUserPos As Position 'Si se mueve
Public UserCharIndex As Integer

Public EngineRun As Boolean

Public FPS As Long
Public FramesPerSecCounter As Long
Private fpsLastCheck As Long

'Tamaño del la vista en Tiles
Private WindowTileWidth As Integer
Private WindowTileHeight As Integer

Private HalfWindowTileWidth As Integer
Private HalfWindowTileHeight As Integer

'Offset del desde 0,0 del main view
Private MainViewTop As Integer
Private MainViewLeft As Integer

'Cuantos tiles el engine mete en el BUFFER cuando
'dibuja el mapa. Ojo un tamaño muy grande puede
'volver el engine muy lento
Public TileBufferSize As Integer

Private TileBufferPixelOffsetX As Integer
Private TileBufferPixelOffsetY As Integer

'Tamaño de los tiles en pixels
Public TilePixelHeight As Integer
Public TilePixelWidth As Integer

'Number of pixels the engine scrolls per frame. MUST divide evenly into pixels per tile
Public ScrollPixelsPerFrameX As Integer
Public ScrollPixelsPerFrameY As Integer

Public timerElapsedTime As Single
Dim timerTicksPerFrame As Double
Public engineBaseSpeed As Single

Public Numheads As Integer
Public NumFxs As Integer

Public NumChars As Integer
Public LastChar As Integer
Public NumWeaponAnims As Integer
Public NumShieldAnims As Integer


Private MainDestRect   As RECT
Private MainViewRect   As RECT
Private BackBufferRect As RECT
Private SetConnect     As RECT

Private MainViewWidth As Integer
Private MainViewHeight As Integer

Private MouseTileX As Byte
Private MouseTileY As Byte

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Graficos¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public GrhData() As GrhData 'Guarda todos los grh
Public BodyData() As BodyData
Public HeadData() As HeadData
Public FxData() As tIndiceFx
Public WeaponAnimData() As WeaponAnimData
Public ShieldAnimData() As ShieldAnimData
Public CascoAnimData() As HeadData
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Mapa?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public MapData() As MapBlock ' Mapa
Public MapInfo As MapInfo ' Info acerca del mapa en uso
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

Public bTecho       As Boolean 'hay techo?
Public bTechoAB As Byte

Public charlist(1 To 10000) As Char

' Used by GetTextExtentPoint32
Private Type Size
    cx As Long
    cy As Long
End Type

Public Enum PlayLoop
    plNone = 0
    plAmbient = 1
End Enum
'
'       [END]
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'Very percise counter 64bit system counter
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

'Text width computation. Needed to center text.
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcsrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal crTransparent As Long) As Long
Public Const PI As Single = 3.14159265358979

'*****************************
Public base_light As Long
Public LightIluminado(3) As Long
Public LightOscurito(3) As Long
Public NoPuedeUsar(3) As Long

Private Sub InitColor()
'*******************************
'By Lorwik
'Iniciamos los colores
'*******************************
Dim i As Long
    bTechoAB = 255
    
    For i = 0 To 3
        LightIluminado(i) = RGB(255, 255, 255)
        LightOscurito(i) = RGB(150, 150, 150)
        NoPuedeUsar(i) = RGB(0, 0, 255)
    Next i
End Sub

Sub ConvertCPtoTP(ByVal viewPortX As Integer, ByVal viewPortY As Integer, ByRef tX As Byte, ByRef tY As Byte)
'******************************************
'Converts where the mouse is in the main window to a tile position. MUST be called eveytime the mouse moves.
'******************************************
    tX = UserPos.X + viewPortX \ TilePixelWidth - WindowTileWidth \ 2
    tY = UserPos.Y + viewPortY \ TilePixelHeight - WindowTileHeight \ 2
End Sub

Sub MakeChar(ByVal CharIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal X As Integer, ByVal Y As Integer, ByVal arma As Integer, ByVal escudo As Integer, ByVal casco As Integer, ByVal Aura_Index As Integer)
On Error Resume Next
    'Apuntamos al ultimo Char
    If CharIndex > LastChar Then LastChar = CharIndex
    
    With charlist(CharIndex)
        'If the char wasn't allready active (we are rewritting it) don't increase char count
        If .Active = 0 Then _
            NumChars = NumChars + 1
        
        If arma = 0 Then arma = 2
        If escudo = 0 Then escudo = 2
        If casco = 0 Then casco = 2
        
        .iHead = Head
        .iBody = Body
        .Head = HeadData(Head)
        .Body = BodyData(Body)
        .arma = WeaponAnimData(arma)
        
        If Aura_Index > 0 Then _
            Call InitGrh(.Aura, Aura_Index)
        
        '[ANIM ATAK]
        .arma.WeaponAttack = 0
                
        .escudo = ShieldAnimData(escudo)
        .casco = CascoAnimData(casco)
        
        .Heading = Heading
        
        'Reset moving stats
        .Moving = 0
        .MoveOffsetX = 0
        .MoveOffsetY = 0
        
        'Update position
        .Pos.X = X
        .Pos.Y = Y
        
        'Make active
        .Active = 1
    End With
    
    'Plot on map
    MapData(X, Y).CharIndex = CharIndex
End Sub

Sub ResetCharInfo(ByVal CharIndex As Integer)
    With charlist(CharIndex)
        .Active = 0
        .Criminal = 0
        .Atacable = False
        Char_Particle_Group_Remove_All (CharIndex)
        .FxIndex = 0
        .invisible = False
        .Moving = 0
        .Muerto = False
        .nombre = ""
        .pie = False
        .Pos.X = 0
        .Pos.Y = 0
        .UsandoArma = False
    End With
End Sub

Sub EraseChar(ByVal CharIndex As Integer)
'*****************************************************************
'Erases a character from CharList and map
'*****************************************************************
On Error Resume Next
    Call Char_Particle_Group_Remove_All(CharIndex)
    charlist(CharIndex).Active = 0
    
    'Update lastchar
    If CharIndex = LastChar Then
        Do Until charlist(LastChar).Active = 1
            LastChar = LastChar - 1
            If LastChar = 0 Then Exit Do
        Loop
    End If
    
    MapData(charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y).CharIndex = 0
    
    'Remove char's dialog
    Call Dialogos.RemoveDialog(CharIndex)
    
    Call ResetCharInfo(CharIndex)
    
    'Update NumChars
    NumChars = NumChars - 1
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
        Grh.Loops = INFINITE_LOOPS
    Else
        Grh.Loops = 0
    End If
    
    Grh.FrameCounter = 1
    Grh.Speed = GrhData(Grh.GrhIndex).Speed
End Sub

Sub MoveCharbyHead(ByVal CharIndex As Integer, ByVal nHeading As E_Heading)
'*****************************************************************
'Starts the movement of a character in nHeading direction
'*****************************************************************
    Dim addx As Integer
    Dim addy As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim nX As Integer
    Dim nY As Integer
    
    With charlist(CharIndex)
        X = .Pos.X
        Y = .Pos.Y
        
        'Figure out which way to move
        Select Case nHeading
            Case E_Heading.NORTH
                addy = -1
        
            Case E_Heading.EAST
                addx = 1
        
            Case E_Heading.SOUTH
                addy = 1
            
            Case E_Heading.WEST
                addx = -1
        End Select
        
        nX = X + addx
        nY = Y + addy
        
        MapData(nX, nY).CharIndex = CharIndex
        .Pos.X = nX
        .Pos.Y = nY
        MapData(X, Y).CharIndex = 0
        
        .MoveOffsetX = -1 * (TilePixelWidth * addx)
        .MoveOffsetY = -1 * (TilePixelHeight * addy)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = addx
        .scrollDirectionY = addy
    End With
    
    If UserEstado = 0 Then Call DoPasosFx(CharIndex)
    
    'areas viejos
    If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
        If CharIndex <> UserCharIndex Then
            Call EraseChar(CharIndex)
        End If
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
        If bFogata And FogataBufferIndex = 0 Then FogataBufferIndex = General_Set_Wav("fuego.wav", location.X, location.Y, LoopStyle.Enabled)
    End If
End Sub

Private Function EstaPCarea(ByVal CharIndex As Integer) As Boolean
    With charlist(CharIndex).Pos
        EstaPCarea = .X > UserPos.X - MinXBorder And .X < UserPos.X + MinXBorder And .Y > UserPos.Y - MinYBorder And .Y < UserPos.Y + MinYBorder
    End With
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
Dim Music As String

    With charlist(CharIndex)
   
    If Not .Muerto And EstaPCarea(CharIndex) Then
        If UserNavegando Then
            If TickON(0, 8) Then General_Set_Wav "50.wav", .Pos.X, .Pos.Y
            Exit Sub
        End If
        
        'En Bosque
        If MapData(.Pos.X, .Pos.Y).Graphic(1).GrhIndex >= 6000 And MapData(.Pos.X, .Pos.Y).Graphic(1).GrhIndex <= 6307 Then
            If Not UserNavegando And .pie Then Music = "237.Wav"
            If Not UserNavegando And Not .pie Then Music = "238.Wav"
                
        'En Nieve
        ElseIf MapData(.Pos.X, .Pos.Y).Graphic(1).GrhIndex >= 22563 And MapData(.Pos.X, .Pos.Y).Graphic(1).GrhIndex <= 22883 Then
            If Not UserNavegando And .pie Then Music = "240.Wav"
            If Not UserNavegando And Not .pie Then Music = "241.Wav"
            
        'En Desierto
        ElseIf MapData(.Pos.X, .Pos.Y).Graphic(1).GrhIndex >= 7704 And MapData(.Pos.X, .Pos.Y).Graphic(1).GrhIndex <= 7719 Then
            If Not UserNavegando And .pie Then Music = "238.Wav"
            If Not UserNavegando And Not .pie Then Music = "239.Wav"
        Else
            If Not UserNavegando And .pie Then Music = "23.wav"
            If Not UserNavegando And Not .pie Then Music = "24.Wav"
        End If
            
        General_Set_Wav Music, .Pos.X, .Pos.Y
        .pie = Not .pie
    End If
    End With
End Sub

Sub MoveCharbyPos(ByVal CharIndex As Integer, ByVal nX As Integer, ByVal nY As Integer)
On Error Resume Next
    Dim X As Integer
    Dim Y As Integer
    Dim addx As Integer
    Dim addy As Integer
    Dim nHeading As E_Heading
    
    With charlist(CharIndex)
        X = .Pos.X
        Y = .Pos.Y
        
        MapData(X, Y).CharIndex = 0
        
        addx = nX - X
        addy = nY - Y
        
        If Sgn(addx) = 1 Then
            nHeading = E_Heading.EAST
        ElseIf Sgn(addx) = -1 Then
            nHeading = E_Heading.WEST
        ElseIf Sgn(addy) = -1 Then
            nHeading = E_Heading.NORTH
        ElseIf Sgn(addy) = 1 Then
            nHeading = E_Heading.SOUTH
        End If
        
        MapData(nX, nY).CharIndex = CharIndex
        
        .Pos.X = nX
        .Pos.Y = nY
        
        .MoveOffsetX = -1 * (TilePixelWidth * addx)
        .MoveOffsetY = -1 * (TilePixelHeight * addy)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = Sgn(addx)
        .scrollDirectionY = Sgn(addy)
    End With
    
    If Not EstaPCarea(CharIndex) Then Call Dialogos.RemoveDialog(CharIndex)
    
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
        
        bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
    End If
End Sub

Private Function HayFogata(ByRef location As Position) As Boolean
    Dim j As Long
    Dim k As Long
    
    For j = UserPos.X - 8 To UserPos.X + 8
        For k = UserPos.Y - 6 To UserPos.Y + 6
            If InMapBounds(j, k) Then
                If MapData(j, k).ObjGrh.GrhIndex = FOgata Then
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
    Dim loopc As Long
    Dim Dale As Boolean
    
    loopc = 1
    Do While charlist(loopc).Active And Dale
        loopc = loopc + 1
        Dale = (loopc <= UBound(charlist))
    Loop
    
    NextOpenChar = loopc
End Function
Function MoveToLegalPos(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Author: Lorwik
'Last Modify Date: 09/01/2011
'******************************************************
    Dim CharIndex As Integer
    
    'Limites del mapa
    If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
        Exit Function
    End If
    
    'Tile Bloqueado?
    If MapData(X, Y).blocked = 1 Then
        Exit Function
    End If
    
    CharIndex = MapData(X, Y).CharIndex
    '¿Hay un personaje?
    If CharIndex > 0 Then
    
        If MapData(UserPos.X, UserPos.Y).blocked = 1 Then
            Exit Function
        End If
        
        With charlist(CharIndex)
            ' Si no es casper, no puede pasar
            If .iHead <> CASPER_HEAD And .iBody <> FRAGATA_FANTASMAL Then
                Exit Function
            Else
                ' No puedo intercambiar con un casper que este en la orilla (Lado tierra)
                If HayAgua(UserPos.X, UserPos.Y) Then
                    If Not HayAgua(X, Y) Then Exit Function
                Else
                    ' No puedo intercambiar con un casper que este en la orilla (Lado agua)
                    If HayAgua(X, Y) Then Exit Function
                End If
                
                ' Los admins no pueden intercambiar pos con caspers cuando estan invisibles
                If charlist(UserCharIndex).priv > 0 And charlist(UserCharIndex).priv < 6 Then
                    If charlist(UserCharIndex).invisible = True Then Exit Function
                End If
            End If
        End With
    End If
   
    If UserNavegando <> HayAgua(X, Y) Then
        Exit Function
    End If
    
        '¿Esta el usuario Equitando bajo un techo?
    If UserEquitando Then
        If bTecho = True Then
            MoveToLegalPos = False
            Exit Function
        End If
    End If
    
    MoveToLegalPos = True
End Function

Function InMapBounds(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************
    If X < XMinMapSize Or X > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
        Exit Function
    End If
    
    InMapBounds = True
End Function

'******************************************************
'======================================================
'*LORWIK> PARA DIBUJAR COSITAS CON DX :)              *
'======================================================
'******************************************************
Private Sub DDrawGrhtoSurface(ByRef Grh As Grh, ByVal X As Integer, ByVal Y As Integer, ByVal center As Byte, ByVal Animate As Byte, ByRef Light() As Long)
    Dim CurrentGrhIndex As Integer
    Dim SourceRect As RECT
On Error GoTo error
        
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed)
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                
                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    With GrhData(CurrentGrhIndex)
        'Center Grh over X,Y pos
        If center Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
        
        SourceRect.Left = .SX
        SourceRect.Top = .SY
        SourceRect.Right = SourceRect.Left + .pixelWidth
        SourceRect.bottom = SourceRect.Top + .pixelHeight
        
        'Draw
        Call Device_Textured_Render(X, Y, SurfaceDB.Surface(.FileNum), SourceRect, Light)
    End With
Exit Sub

error:
    If Err.number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        MsgBox "Ocurrió un error inesperado, por favor comuniquelo a los administradores del juego." & vbCrLf & "Descripción del error: " & _
        vbCrLf & Err.Description, vbExclamation, "[ " & Err.number & " ] Error"
        End
    End If
End Sub

Sub DDrawTransGrhIndextoSurface(ByVal GrhIndex As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal center As Byte, ByRef Light() As Long)
    Dim SourceRect As RECT
    
    With GrhData(GrhIndex)
        'Center Grh over X,Y pos
        If center Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
        
        SourceRect.Left = .SX
        SourceRect.Top = .SY
        SourceRect.Right = SourceRect.Left + .pixelWidth
        SourceRect.bottom = SourceRect.Top + .pixelHeight
        
        'Draw
        'Call BackBufferSurface.BltFast(X, Y, SurfaceDB.Surface(.FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
        Call Device_Textured_Render(X, Y, SurfaceDB.Surface(.FileNum), SourceRect, Light)
    End With
End Sub

Sub DDrawTransGrhtoSurface(ByRef Grh As Grh, ByVal X As Integer, ByVal Y As Integer, ByVal center As Byte, ByVal Animate As Byte, ByRef Light() As Long, Optional Transp As Byte = 255, Optional blend As Boolean, Optional angle As Single)
'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
    Dim CurrentGrhIndex As Integer
    Dim SourceRect As RECT
'On Error GoTo error
   If Grh.GrhIndex = 0 Then Exit Sub
   
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed) * movSpeed
            
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                
                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    With GrhData(CurrentGrhIndex)
        'Center Grh over X,Y pos
        If center Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
                
        SourceRect.Left = .SX
        SourceRect.Top = .SY
        SourceRect.Right = SourceRect.Left + .pixelWidth
        SourceRect.bottom = SourceRect.Top + .pixelHeight
        
        'Draw
                
        Call Device_Textured_Render(X, Y, SurfaceDB.Surface(.FileNum), SourceRect, Light(), CBool(blend), Transp, angle)
    End With
    
    
Exit Sub

error:
    If Err.number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        MsgBox "Ocurrió un error inesperado, por favor comuniquelo a los administradores del juego." & vbCrLf & "Descripción del error: " & _
        vbCrLf & Err.Description, vbExclamation, "[ " & Err.number & " ] Error"
        End
    End If
End Sub
'******************************************************
'======================================================
'*/LORWIK> PARA DIBUJAR COSITAS CON DX :)             *
'======================================================
'******************************************************

Function GetBitmapDimensions(ByVal BmpFile As String, ByRef bmWidth As Long, ByRef bmHeight As Long)
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

Public Sub DrawGrhtoHdc(desthdc As Long, ByVal grh_index As Long, ByVal screen_x As Long, ByVal screen_y As Long, Optional transparent As Boolean = False)

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
    Dim InfoHead As INFOHEADER
        
    InfoHead = File_Find(App.Path & "\RECURSOS\Graphics.WAO", CStr(GrhData(grh_index).FileNum) & ".bmp")
    
    If grh_index <= 0 Then Exit Sub
    
    'If it's animated switch grh_index to first frame
    If GrhData(grh_index).NumFrames <> 1 Then
    grh_index = GrhData(grh_index).Frames(1)
    End If
        
    file_path = Get_Extract(graphics, CStr(GrhData(grh_index).FileNum) & ".bmp")
        
    Src_X = GrhData(grh_index).SX
    Src_Y = GrhData(grh_index).SY
    src_width = GrhData(grh_index).pixelWidth
    src_height = GrhData(grh_index).pixelHeight
    
    hdcsrc = CreateCompatibleDC(desthdc)
    
    PrevObj = SelectObject(hdcsrc, LoadPicture(file_path))
    
    If transparent = False Then
        BitBlt desthdc, screen_x, screen_y, src_width, src_height, hdcsrc, Src_X, Src_Y, vbSrcCopy
    Else
        TransparentBlt desthdc, screen_x, screen_y, src_width, src_height, hdcsrc, Src_X, Src_Y, src_width, src_height, &HFF000000
    End If
    
    DeleteDC hdcsrc
    Delete_File file_path

End Sub


Public Sub DrawImageInPicture(ByRef PictureBox As PictureBox, ByRef Picture As StdPicture, ByVal x1 As Single, ByVal y1 As Single, Optional Width1, Optional Height1, Optional x2, Optional y2, Optional Width2, Optional Height2)
'**************************************************************
'Author: Torres Patricio (Pato)
'Last Modify Date: 12/28/2009
'Draw Picture in the PictureBox
'*************************************************************
Call PictureBox.PaintPicture(Picture, x1, y1, Width1, Height1, x2, y2, Width2, Height2)
End Sub


Sub RenderScreen(ByVal tilex As Integer, ByVal tiley As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 8/14/2007
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Renders everything to the viewport
'**************************************************************
    Dim Y           As Long     'Keeps track of where on map we are
    Dim X           As Long     'Keeps track of where on map we are
    Dim screenminY  As Integer  'Start Y pos on current screen
    Dim screenmaxY  As Integer  'End Y pos on current screen
    Dim screenminX  As Integer  'Start X pos on current screen
    Dim screenmaxX  As Integer  'End X pos on current screen
    Dim minY        As Integer  'Start Y pos on current map
    Dim maxY        As Integer  'End Y pos on current map
    Dim minX        As Integer  'Start X pos on current map
    Dim maxX        As Integer  'End X pos on current map
    Dim ScreenX     As Integer  'Keeps track of where to place tile on screen
    Dim ScreenY     As Integer  'Keeps track of where to place tile on screen
    Dim minXOffset  As Integer
    Dim minYOffset  As Integer
    Dim PixelOffsetXTemp As Integer 'For centering grhs
    Dim PixelOffsetYTemp As Integer 'For centering grhs
    
    If TwinkLightByteHandle = 15 Then
        TwinkLightByteHandle = 0
    Else
        TwinkLightByteHandle = TwinkLightByteHandle + timerTicksPerFrame * 5
        If TwinkLightByteHandle > 15 Then TwinkLightByteHandle = 15
    End If
            
    'Figure out Ends and Starts of screen
    screenminY = tiley - HalfWindowTileHeight
    screenmaxY = tiley + HalfWindowTileHeight
    screenminX = tilex - HalfWindowTileWidth
    screenmaxX = tilex + HalfWindowTileWidth
    
    minY = screenminY - TileBufferSize
    maxY = screenmaxY + TileBufferSize
    minX = screenminX - TileBufferSize
    maxX = screenmaxX + TileBufferSize
        
    'Make sure mins and maxs are allways in map bounds
    If minY < XMinMapSize Then
        minYOffset = YMinMapSize - minY
        minY = YMinMapSize
    End If
    
    If maxY > YMaxMapSize Then maxY = YMaxMapSize
    
    If minX < XMinMapSize Then
        minXOffset = XMinMapSize - minX
        minX = XMinMapSize
    End If
    
    If maxX > XMaxMapSize Then maxX = XMaxMapSize
    
    'If we can, we render around the view area to make it smoother
    If screenminY > YMinMapSize Then
        screenminY = screenminY - 1
    Else
        screenminY = 1
        ScreenY = 1
    End If
    
    If screenmaxY < YMaxMapSize Then screenmaxY = screenmaxY + 1
    
    If screenminX > XMinMapSize Then
        screenminX = screenminX - 1
    Else
        screenminX = 1
        ScreenX = 1
    End If
    
    If screenmaxX < XMaxMapSize Then screenmaxX = screenmaxX + 1
        
    'Draw floor layer
    For Y = screenminY To screenmaxY
        For X = screenminX To screenmaxX
            'Layer 1 **********************************
            If MapData(X, Y).Graphic(1).GrhIndex <> 0 Then _
                Call DDrawGrhtoSurface(MapData(X, Y).Graphic(1), (ScreenX - 1) * TilePixelWidth + PixelOffsetX, (ScreenY - 1) * TilePixelHeight + PixelOffsetY, 0, 1, MapData(X, Y).light_value)
            '******************************************
            
            ScreenX = ScreenX + 1
        Next X
        ScreenX = ScreenX - X + screenminX
        ScreenY = ScreenY + 1
    Next Y
    
     'Draw floor layer 2
    ScreenY = minYOffset - TileBufferSize
    For Y = minY To maxY
        ScreenX = minXOffset - TileBufferSize
        For X = minX To maxX
            'Layer 2 **********************************
            If MapData(X, Y).Graphic(2).GrhIndex <> 0 Then _
                Call DDrawTransGrhtoSurface(MapData(X, Y).Graphic(2), ScreenX * TilePixelWidth + PixelOffsetX, ScreenY * TilePixelHeight + PixelOffsetY, 1, 1, MapData(X, Y).light_value)
            '******************************************
                ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y
    
    'Draw Transparent Layers
    ScreenY = minYOffset - TileBufferSize
    For Y = minY To maxY
        ScreenX = minXOffset - TileBufferSize
        For X = minX To maxX
            PixelOffsetXTemp = ScreenX * TilePixelWidth + PixelOffsetX
            PixelOffsetYTemp = ScreenY * TilePixelHeight + PixelOffsetY
            With MapData(X, Y)
            
                'Sangre **********************************
                If .Blood.Active = 1 And Opciones.SangreAct Then
                    If .Blood.LifeTime >= 0 Then .Blood.LifeTime = .Blood.LifeTime - 1
                    If .Blood.LifeTime <= 0 Then .Blood.Alpha = .Blood.Alpha - 1
                                        
                    Dim i As Byte
                    For i = 0 To 3
                        .Blood.color(i) = D3DColorARGB(.Blood.Alpha, 255, 0, 0)
                    Next i
                    
                    If .Blood.Alpha <= 5 Then _
                        .Blood.Active = 0
                                                
                    Select Case .Blood.Head
                        Case E_Heading.EAST 'derecha
                            Call DDrawTransGrhtoSurface(.Blood.Grh, PixelOffsetXTemp - 20, PixelOffsetYTemp, 1, 1, .Blood.color())
                        Case E_Heading.NORTH 'arriba
                            Call DDrawTransGrhtoSurface(.Blood.Grh, PixelOffsetXTemp, PixelOffsetYTemp + 20, 1, 1, .Blood.color())
                        Case E_Heading.SOUTH 'abajo
                            Call DDrawTransGrhtoSurface(.Blood.Grh, PixelOffsetXTemp, PixelOffsetYTemp - 20, 1, 1, .Blood.color())
                        Case E_Heading.WEST 'izquierda
                           Call DDrawTransGrhtoSurface(.Blood.Grh, PixelOffsetXTemp + 20, PixelOffsetYTemp, 1, 1, .Blood.color())
                        End Select
                End If
                '**********************************
                
                'Object Layer **********************************
                If .ObjGrh.GrhIndex <> 0 Then
                        Call DDrawTransGrhtoSurface(.ObjGrh, PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, .light_value)
                End If
                '***********************************************
                
                'Char layer ************************************
                If .CharIndex <> 0 Then
                    Call CharRender(.CharIndex, PixelOffsetXTemp, PixelOffsetYTemp, .light_value)
                End If
                '*************************************************
                
                'Layer 3 *****************************************
                If .Graphic(3).GrhIndex <> 0 Then
                    'Draw
                    Call DDrawTransGrhtoSurface(.Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, .light_value)
                End If
                '************************************************
                
            End With
            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y
    
    'Capa 3.5 (Particulas)
    ScreenY = minYOffset - TileBufferSize
        For Y = minY To maxY
            ScreenX = minXOffset - TileBufferSize
                For X = minX To maxX
                    'Particulas**************************************
                    If MapData(X, Y).particle_group_index Then
                        Particle_Group_Render MapData(X, Y).particle_group_index, ScreenX * 32 + PixelOffsetX, ScreenY * 32 + PixelOffsetY
                    '************************************************
                    End If
                ScreenX = ScreenX + 1
            Next X
        ScreenY = ScreenY + 1
    Next Y


    If Not bTecho And bTechoAB < 255 Then
        bTechoAB = bTechoAB + 1
    ElseIf bTecho And bTechoAB > 0 Then
        bTechoAB = bTechoAB - 1
    End If
    
    'Draw blocked tiles and grid
    ScreenY = minYOffset - TileBufferSize
    For Y = minY To maxY
        ScreenX = minXOffset - TileBufferSize
        For X = minX To maxX
            'Layer 4 **********************************
            If MapData(X, Y).Graphic(4).GrhIndex Then
                'Draw
                Call DDrawTransGrhtoSurface(MapData(X, Y).Graphic(4), ScreenX * TilePixelWidth + PixelOffsetX, ScreenY * TilePixelHeight + PixelOffsetY, 1, 1, MapData(X, Y).light_value, bTechoAB)
            End If
            '**********************************
            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y

    Call ClimaX
End Sub

Public Function RenderSounds()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/30/2008
'Actualiza todos los sonidos del mapa.
'**************************************************************
    Call Audio.MusicMP3GetLoop
    DoFogataFx
End Function
Public Sub Ambient()
'**************************************************************
'Author: Manuel (Lorwik)
'Last Modify Date: 6/18/2011
'Actualiza todos los sonidos del mapa.
'**************************************************************
If Opciones.AmbientAct = True Then
    
    'Si es de noche reproducimos el Ambient de los grillos
    If Not Zona = "DUNGEON" Then
        If Anocheceria = 3 Then
            Call Audio.StopWave
            Call General_Set_Wav("230.wav", , , Enabled)
            frmMain.IsPlaying = PlayLoop.plAmbient
            Exit Sub
        End If
    End If
    
    Dim Ambient As String
    Dim file As String
        'Lorwik> Sistema Ambient chapucero (Asi se queda de momento *Yao Ming*)
         file = Get_Extract(Scripts, "WorldMapData.dat")
         Ambient = GetVar(file, "AMBIENT", UserMap)
         If Not Ambient = "" Then
            Call Audio.StopWave
            Call General_Set_Wav(Ambient & ".wav", , , Enabled)
            frmMain.IsPlaying = PlayLoop.plAmbient
         Else
            Call Audio.StopWave
         End If
         Delete_File file
End If
End Sub
Public Function InitTileEngine(ByVal setDisplayFormhWnd As Long, ByVal setMainViewTop As Integer, ByVal setMainViewLeft As Integer, ByVal setTilePixelHeight As Integer, ByVal setTilePixelWidth As Integer, ByVal setWindowTileHeight As Integer, ByVal setWindowTileWidth As Integer, ByVal setTileBufferSize As Integer, ByVal pixelsToScrollPerFrameX As Integer, pixelsToScrollPerFrameY As Integer, ByVal engineSpeed As Single) As Boolean
    
    'Fill startup variables
    MainViewTop = setMainViewTop
    MainViewLeft = setMainViewLeft
    TilePixelWidth = setTilePixelWidth
    TilePixelHeight = setTilePixelHeight
    WindowTileHeight = setWindowTileHeight
    WindowTileWidth = setWindowTileWidth
    TileBufferSize = setTileBufferSize
    
    HalfWindowTileHeight = setWindowTileHeight \ 2
    HalfWindowTileWidth = setWindowTileWidth \ 2
    
    'Compute offset in pixels when rendering tile buffer.
    'We diminish by one to get the top-left corner of the tile for rendering.
    TileBufferPixelOffsetX = ((TileBufferSize - 1) * TilePixelWidth)
    TileBufferPixelOffsetY = ((TileBufferSize - 1) * TilePixelHeight)
    
    engineBaseSpeed = engineSpeed
    
    MinXBorder = XMinMapSize + (WindowTileWidth \ 2)
    MaxXBorder = XMaxMapSize - (WindowTileWidth \ 2)
    MinYBorder = YMinMapSize + (WindowTileHeight \ 2)
    MaxYBorder = YMaxMapSize - (WindowTileHeight \ 2)
    
    MainViewWidth = TilePixelWidth * WindowTileWidth
    MainViewHeight = TilePixelHeight * WindowTileHeight
    
    'Resize mapdata array
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    
    'Set intial user position
    UserPos.X = MinXBorder
    UserPos.Y = MinYBorder
    
    'Set scroll pixels per frame
    ScrollPixelsPerFrameX = pixelsToScrollPerFrameX
    ScrollPixelsPerFrameY = pixelsToScrollPerFrameY
    
    'Set the dest rect
    With MainDestRect
        .Left = TilePixelWidth * TileBufferSize - TilePixelWidth
        .Top = TilePixelHeight * TileBufferSize - TilePixelHeight
        .Right = .Left + MainViewWidth
        .bottom = .Top + MainViewHeight
    End With
    
On Error GoTo 0
    
    Call LoadGrhData
    Call CargarCuerpos
    Call CargarCabezas
    Call CargarCascos
    Call CargarFxs
    CargarParticulas
    movSpeed = 1
    InitColor
    
    Call SurfaceDB.Initialize(DirectD3D8, ClientSetup.bUseVideo, DirGraficos, 40)
    
    InitTileEngine = True
End Function
Public Sub DirectXInit()
    Dim DispMode As D3DDISPLAYMODE
    Dim D3DWindow As D3DPRESENT_PARAMETERS
    
    Set DirectX = New DirectX8
    Set DirectD3D = DirectX.Direct3DCreate
    Set DirectD3D8 = New D3DX8
    
    DirectD3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
    
    Opciones.VSync = GetVar(App.Path & "\Init\Config.cfg", "Video", "VSync")
    
    With D3DWindow
        .Windowed = True
        If Opciones.VSync = True Then
            .SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
        Else
            .SwapEffect = D3DSWAPEFFECT_COPY
        End If
        .BackBufferFormat = DispMode.Format
        .BackBufferWidth = frmMain.MainViewPic.ScaleWidth
        .BackBufferHeight = frmMain.MainViewPic.ScaleHeight
        .EnableAutoDepthStencil = 1
        .AutoDepthStencilFormat = D3DFMT_D16
        .hDeviceWindow = frmMain.MainViewPic.hwnd
    End With

    Set DirectDevice = DirectD3D.CreateDevice( _
                        D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, _
                        frmMain.MainViewPic.hwnd, _
                        D3DCREATE_SOFTWARE_VERTEXPROCESSING, _
                        D3DWindow)

    DirectDevice.SetVertexShader D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR
    
    With DirectDevice
        .SetRenderState D3DRS_LIGHTING, False
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
        .SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
        .SetTextureStageState 0, D3DTSS_ALPHAARG2, D3DTA_TFACTOR
    End With
    
    If Err Then
        MsgBox "No se puede iniciar DirectX. Por favor asegurese de tener la ultima version correctamente instalada."
        Exit Sub
    End If
    
    If Err Then
        MsgBox "No se puede iniciar DirectD3D. Por favor asegurese de tener la ultima version correctamente instalada."
        Exit Sub
    End If
    
    If DirectDevice Is Nothing Then
        MsgBox "No se puede inicializar DirectDevice. Por favor asegurese de tener la ultima version correctamente instalada."
        Exit Sub
    End If
End Sub
Public Sub DeinitTileEngine()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/14/07
'Destroys all DX objects
'***************************************************
On Error Resume Next

    Set DirectD3D = Nothing
    
    Set DirectX = Nothing
End Sub

Sub ShowNextFrame(ByVal DisplayFormTop As Integer, ByVal DisplayFormLeft As Integer, ByVal MouseViewX As Integer, ByVal MouseViewY As Integer)
'***************************************************
'Author: Arron Perkins
'Last Modification: 08/14/07
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Updates the game's model and renders everything.
'***************************************************
    Static OffsetCounterX As Single
    Static OffsetCounterY As Single
    
    '****** Set main view rectangle ******
    MainViewRect.Left = (DisplayFormLeft / Screen.TwipsPerPixelX) + MainViewLeft
    MainViewRect.Top = (DisplayFormTop / Screen.TwipsPerPixelY) + MainViewTop
    MainViewRect.Right = MainViewRect.Left + MainViewWidth
    MainViewRect.bottom = MainViewRect.Top + MainViewHeight
    
    If EngineRun Then
        DirectDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
        DirectDevice.BeginScene
        
        If UserMoving Then
            '****** Move screen Left and Right if needed ******
            If AddtoUserPos.X <> 0 Then
                OffsetCounterX = OffsetCounterX - ScrollPixelsPerFrameX * AddtoUserPos.X * timerTicksPerFrame
                If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.X) Then
                    OffsetCounterX = 0
                    AddtoUserPos.X = 0
                    UserMoving = False
                End If
            End If
            
            '****** Move screen Up and Down if needed ******
            If AddtoUserPos.Y <> 0 Then
                OffsetCounterY = OffsetCounterY - ScrollPixelsPerFrameY * AddtoUserPos.Y * timerTicksPerFrame
                If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.Y) Then
                    OffsetCounterY = 0
                    AddtoUserPos.Y = 0
                    UserMoving = False
                End If
            End If
        End If
        
        'Update mouse position within view area
        Call ConvertCPtoTP(MouseViewX, MouseViewY, MouseTileX, MouseTileY)
        
        '****** Update screen ******
        If Not UserCiego Then
            Call RenderScreen(UserPos.X - AddtoUserPos.X, UserPos.Y - AddtoUserPos.Y, OffsetCounterX, OffsetCounterY)
        End If

        '////////////CARTELES\\\\\\\\\\\\\\\
        If IScombate = True Then Fonts_Render_String "MODO COMBATE", 1, 1, vbBlue
        
        '*********Tiempo restante para que termine el invi o el paralizar*********
        If CartelInvisibilidad > 0 Then Fonts_Render_String CartelInvisibilidad & " segundos restantes de Invisibilidad", 1, 13, vbCyan
        If CartelParalisis > 0 Then Fonts_Render_String CartelParalisis & " segundos restantes de Paralisis", 1, 25, vbGreen
        '*************************************************************************
        '|||||||||||||||||||||||||||||||||||
        
        Call Dialogos.Render
        Call DibujarCartel
        
        'FPS update
        If fpsLastCheck + 1000 < GetTickCount Then
            'FPS = FramesPerSecCounter
            FramesPerSecCounter = 1
            fpsLastCheck = GetTickCount
        Else
            FramesPerSecCounter = FramesPerSecCounter + 1
        End If
    
        'Get timing info
        timerElapsedTime = GetElapsedTime()
        timerTicksPerFrame = timerElapsedTime * engineBaseSpeed
        FPS = 1000 / timerElapsedTime
        
        DirectDevice.EndScene
        DirectDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
    End If
End Sub

Private Function GetElapsedTime() As Single
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the time that past since the last call
'**************************************************************
    Dim start_time As Currency
    Static end_time As Currency
    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        QueryPerformanceFrequency timer_freq
    End If
    
    'Get current time
    Call QueryPerformanceCounter(start_time)
    
    'Calculate elapsed time
    GetElapsedTime = (start_time - end_time) / timer_freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(end_time)
End Function

Private Sub CharRender(ByVal CharIndex As Long, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer, Light() As Long, Optional angle As Single)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Draw char's to screen without offcentering them
'***************************************************
    Dim moved As Boolean
    Dim Pos As Integer
    Dim line As String
    Dim color As Long

    With charlist(CharIndex)
    
        If .Moving Then
            '****** Move Left and Right if needed ******
            If .scrollDirectionX <> 0 Then
                .MoveOffsetX = .MoveOffsetX + ScrollPixelsPerFrameX * Sgn(.scrollDirectionX) * timerTicksPerFrame
                
                'Start animations
                If .Body.Walk(.Heading).Speed > 0 Then _
                .Body.Walk(.Heading).Started = 1
                .arma.WeaponWalk(.Heading).Started = 1
                .escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionX) = 1 And .MoveOffsetX >= 0) Or _
                        (Sgn(.scrollDirectionX) = -1 And .MoveOffsetX <= 0) Then
                    .MoveOffsetX = 0
                    .scrollDirectionX = 0
                End If
            End If
            
            '****** Move Up and Down if needed ******
            If .scrollDirectionY <> 0 Then
                .MoveOffsetY = .MoveOffsetY + ScrollPixelsPerFrameY * Sgn(.scrollDirectionY) * timerTicksPerFrame
                
                'Start animations
                If .Body.Walk(.Heading).Speed > 0 Then _
                .Body.Walk(.Heading).Started = 1
                .arma.WeaponWalk(.Heading).Started = 1
                .escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionY) = 1 And .MoveOffsetY >= 0) Or _
                        (Sgn(.scrollDirectionY) = -1 And .MoveOffsetY <= 0) Then
                    .MoveOffsetY = 0
                    .scrollDirectionY = 0
                End If
            End If
        End If
        'End scrolling if needed
        
        'If done moving stop animation
        If Not moved Then
            'Stop animations
            .Body.Walk(.Heading).Started = 0
            .Body.Walk(.Heading).FrameCounter = 1
                
            If IsAttacking = False Then
                .arma.WeaponWalk(.Heading).Started = 0
                .arma.WeaponWalk(.Heading).FrameCounter = 1
                
                .escudo.ShieldWalk(.Heading).Started = 0
                .escudo.ShieldWalk(.Heading).FrameCounter = 1
            End If
                
            .Moving = False
        Else
            IsAttacking = False
        End If
        
        PixelOffsetX = PixelOffsetX + .MoveOffsetX
        PixelOffsetY = PixelOffsetY + .MoveOffsetY
        
        '************Char Normal************
        If .Head.Head(.Heading).GrhIndex Then
            If Not .invisible Then
                movSpeed = 0.5
                
                If .Aura_Index > 0 Then _
                    Call DDrawTransGrhtoSurface(.Aura, PixelOffsetX - 2, PixelOffsetY + 25, 1, 0, Light, 100, 1, angle)
                If .Body.Walk(.Heading).GrhIndex Then _
                    Call DDrawTransGrhtoSurface(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, Light, , , angle)
                If .Head.Head(.Heading).GrhIndex Then
                    Call DDrawTransGrhtoSurface(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y, 1, 0, Light, , , angle)
                If .casco.Head(.Heading).GrhIndex Then _
                    Call DDrawTransGrhtoSurface(.casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y, 1, 0, Light, , , angle)
                If Not UserEquitando Then
                    If .arma.WeaponWalk(.Heading).GrhIndex Then _
                        Call DDrawTransGrhtoSurface(.arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, Light, , , angle)
                    If .escudo.ShieldWalk(.Heading).GrhIndex Then _
                        Call DDrawTransGrhtoSurface(.escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, Light, , , angle)
                End If
                
                    '************Draw name over head************
                    If LenB(.nombre) > 0 Then
                        If Nombres Then
                            Pos = getTagPosition(.nombre)
                            
                            If .priv = 0 Then
                                If .Atacable Then
                                    color = D3DColorXRGB(ColoresPJ(48).r, ColoresPJ(48).g, ColoresPJ(48).B)
                                Else
                                    If .Criminal Then
                                        color = D3DColorXRGB(ColoresPJ(50).r, ColoresPJ(50).g, ColoresPJ(50).B)
                                    Else
                                        color = D3DColorXRGB(ColoresPJ(49).r, ColoresPJ(49).g, ColoresPJ(49).B)
                                    End If
                                End If
                            Else
                                color = D3DColorXRGB(ColoresPJ(.priv).r, ColoresPJ(.priv).g, ColoresPJ(.priv).B)
                            End If
                            
                            '************Nick************
                            line = Left$(.nombre, Pos - 2)
                            Fonts_Render_String line, PixelOffsetX - (Len(line) / 2) * 6 + 15, PixelOffsetY + 30, color
                            '************Clan************
                            line = mid$(.nombre, Pos)
                            Fonts_Render_String line, PixelOffsetX - (Len(line) / 2) * 6 + IIf(Len(line) > 20, 0, 10), PixelOffsetY + 42, D3DColorXRGB(231, 202, 157)
                        End If
                    End If
                End If
            Else
            '************Char Invisible************
                If CharIndex = UserCharIndex Then
                    movSpeed = 0.5
                    
                    If .Aura_Index > 0 Then _
                        Call DDrawTransGrhtoSurface(.Aura, PixelOffsetX - 2, PixelOffsetY + 25, 1, 0, Light, 100, 1, angle)
                    If .Body.Walk(.Heading).GrhIndex Then _
                        Call DDrawTransGrhtoSurface(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, Light, , 1, angle)
                    If .Head.Head(.Heading).GrhIndex Then _
                        Call DDrawTransGrhtoSurface(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y, 1, 0, Light, , 1, angle)
                    If .casco.Head(.Heading).GrhIndex Then _
                        Call DDrawTransGrhtoSurface(.casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y, 1, 0, Light, , 1, angle)
                    If Not UserEquitando Then
                        If .arma.WeaponWalk(.Heading).GrhIndex Then _
                            Call DDrawTransGrhtoSurface(.arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, Light, , 1, angle)
                        If .escudo.ShieldWalk(.Heading).GrhIndex Then _
                            Call DDrawTransGrhtoSurface(.escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, Light, , 1, angle)
                    End If
                End If
            End If
        Else
        '************Si no tiene cabeza mostramos igualmente el nombre************
            If .Body.Walk(.Heading).GrhIndex Then _
                Call DDrawTransGrhtoSurface(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, Light, , , angle)
        End If
        
        '************Update dialogs************
        Call Dialogos.UpdateDialogPos(PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y, CharIndex)    '34 son los pixeles del grh de la cabeza que quedan superpuestos al cuerpo
         movSpeed = 1
         
        '************Particulas************
        Dim i As Integer
            If .particle_count > 0 Then
                For i = 1 To .particle_count
                    If .particle_group(i) > 0 Then _
                        Particle_Group_Render .particle_group(i), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY
                Next i
            End If
         
        '************Draw FX************
        If .FxIndex <> 0 Then
            Call DDrawTransGrhtoSurface(.fX, PixelOffsetX + FxData(.FxIndex).OffsetX, PixelOffsetY + FxData(.FxIndex).OffsetY + 10, 1, 1, Light, 255, 1)
            'Check if animation is over
            If .fX.Started = 0 Then _
                .FxIndex = 0
        End If
        
        '************Draw Pasos************
        If CharIndex = UserCharIndex Then
            If Not UserEquitando Then
                If MapData(.Pos.X, .Pos.Y).Graphic(1).GrhIndex >= 22563 And MapData(.Pos.X, .Pos.Y).Graphic(1).GrhIndex <= 22883 Or MapData(.Pos.X, .Pos.Y).Graphic(1).GrhIndex >= 7704 And MapData(.Pos.X, .Pos.Y).Graphic(1).GrhIndex <= 7719 Then
                    If .Heading = WEST Then
                        Call General_Particle_Create(19, .Pos.X, .Pos.Y, 250)
                    ElseIf .Heading = EAST Then
                        Call General_Particle_Create(17, .Pos.X, .Pos.Y, 250)
                    ElseIf .Heading = NORTH Then
                        Call General_Particle_Create(18, .Pos.X, .Pos.Y, 250)
                    ElseIf .Heading = SOUTH Then
                        Call General_Particle_Create(31, .Pos.X, .Pos.Y, 250)
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub SetCharacterFx(ByVal CharIndex As Integer, ByVal fX As Integer, ByVal Loops As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Sets an FX to the character.
'***************************************************
    With charlist(CharIndex)
        .FxIndex = fX
        
        If .FxIndex > 0 Then
            Call InitGrh(.fX, FxData(fX).Animacion)
        
            .fX.Loops = Loops
        End If
    End With
End Sub

Public Sub Geometry_Create_Box(ByRef verts() As TLVERTEX, ByRef dest As RECT, ByRef src As RECT, ByRef rgb_list() As Long, _
                                Optional ByRef Textures_Width As Long, Optional ByRef Textures_Height As Long, Optional ByVal angle As Single)
'**************************************************************
'Author: Aaron Perkins
'Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 11/17/2002
'
' * v1      * v3
' |\        |
' |  \      |
' |    \    |
' |      \  |
' |        \|
' * v0      * v2
'**************************************************************
    Dim x_center As Single
    Dim y_center As Single
    Dim radius As Single
    Dim x_Cor As Single
    Dim y_Cor As Single
    Dim left_point As Single
    Dim right_point As Single
    Dim temp As Single
    
    If angle > 0 Then
        'Center coordinates on screen of the square
        x_center = dest.Left + (dest.Right - dest.Left) / 2
        y_center = dest.Top + (dest.bottom - dest.Top) / 2
        
        'Calculate radius
        radius = Sqr((dest.Right - x_center) ^ 2 + (dest.bottom - y_center) ^ 2)
        
        'Calculate left and right points
        temp = (dest.Right - x_center) / radius
        right_point = Atn(temp / Sqr(-temp * temp + 1))
        left_point = 3.1459 - right_point
    End If
    
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Left
        y_Cor = dest.bottom
    Else
        x_Cor = x_center + Cos(-left_point - angle) * radius
        y_Cor = y_center - Sin(-left_point - angle) * radius
    End If
    
    
    '0 - Bottom left vertex
    If Textures_Width Or Textures_Height Then
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, src.Left / Textures_Width + 0.001, (src.bottom + 1) / Textures_Height)
    Else
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, 0, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Left
        y_Cor = dest.Top
    Else
        x_Cor = x_center + Cos(left_point - angle) * radius
        y_Cor = y_center - Sin(left_point - angle) * radius
    End If
    
    
    '1 - Top left vertex
    If Textures_Width Or Textures_Height Then
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, src.Left / Textures_Width + 0.001, src.Top / Textures_Height + 0.001)
    Else
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, 0, 1)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.bottom
    Else
        x_Cor = x_center + Cos(-right_point - angle) * radius
        y_Cor = y_center - Sin(-right_point - angle) * radius
    End If
    
    
    '2 - Bottom right vertex
    If Textures_Width Or Textures_Height Then
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, (src.Right + 1) / Textures_Width, (src.bottom + 1) / Textures_Height)
    Else
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, 1, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.Top
    Else
        x_Cor = x_center + Cos(right_point - angle) * radius
        y_Cor = y_center - Sin(right_point - angle) * radius
    End If
    
    
    '3 - Top right vertex
    If Textures_Width Or Textures_Height Then
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, (src.Right + 1) / Textures_Width, src.Top / Textures_Height + 0.001)
    Else
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, 1, 1)
    End If

End Sub
Public Function Geometry_Create_TLVertex(ByVal X As Single, ByVal Y As Single, ByVal z As Single, _
                                            ByVal rhw As Single, ByVal color As Long, ByVal Specular As Long, tu As Single, _
                                            ByVal tv As Single) As TLVERTEX
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'**************************************************************
    Geometry_Create_TLVertex.X = X
    Geometry_Create_TLVertex.Y = Y
    Geometry_Create_TLVertex.z = z
    Geometry_Create_TLVertex.rhw = rhw
    Geometry_Create_TLVertex.color = color
    Geometry_Create_TLVertex.Specular = Specular
    Geometry_Create_TLVertex.tu = tu
    Geometry_Create_TLVertex.tv = tv
End Function
Public Sub Device_Textured_Render(ByVal X As Integer, ByVal Y As Integer, ByVal Texture As Direct3DTexture8, ByRef src_rect As RECT, ByRef rgb_list() As Long, Optional Alpha As Boolean = False, Optional alphabyte As Byte = 255, Optional angle As Single)
    Dim dest_rect As RECT
    Dim temp_verts(3) As TLVERTEX
    Dim srdesc As D3DSURFACE_DESC
    Static light_value(0 To 3) As Long
    
    light_value(0) = rgb_list(0)
    light_value(1) = rgb_list(1)
    light_value(2) = rgb_list(2)
    light_value(3) = rgb_list(3)
    
    If (light_value(0) = 0) Then light_value(0) = base_light
    If (light_value(1) = 0) Then light_value(1) = base_light
    If (light_value(2) = 0) Then light_value(2) = base_light
    If (light_value(3) = 0) Then light_value(3) = base_light
 
    With dest_rect
        .bottom = Y + (src_rect.bottom - src_rect.Top)
        .Left = X
        .Right = X + (src_rect.Right - src_rect.Left)
        .Top = Y
    End With
    
    Dim texwidth As Long, texheight As Long
    Texture.GetLevelDesc 0, srdesc
    texwidth = srdesc.Width
    texheight = srdesc.Height
    
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, light_value(), texwidth, texheight, angle
    
    DirectDevice.SetTexture 0, Texture
    
    If Alpha Then
        DirectDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If
    
    DirectDevice.SetRenderState D3DRS_TEXTUREFACTOR, D3DColorARGB(alphabyte, 0, 0, 0)
    DirectDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
    
    If Alpha Then
        DirectDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If
    
End Sub
Public Sub Draw_FillBox(ByVal X As Integer, ByVal Y As Integer, ByVal Width As Integer, ByVal Height As Integer, color As Long, outlinecolor As Long)

    Static box_rect As RECT
    Static Outline As RECT
    Static rgb_list(3) As Long
    Static rgb_list2(3) As Long
    Static vertex(3) As TLVERTEX
    Static Vertex2(3) As TLVERTEX
    
    rgb_list(0) = color
    rgb_list(1) = color
    rgb_list(2) = color
    rgb_list(3) = color
    
    rgb_list2(0) = outlinecolor
    rgb_list2(1) = outlinecolor
    rgb_list2(2) = outlinecolor
    rgb_list2(3) = outlinecolor
    
    With box_rect
        .bottom = Y + Height
        .Left = X
        .Right = X + Width
        .Top = Y
    End With
    
    With Outline
        .bottom = Y + Height + 1
        .Left = X
        .Right = X + Width + 1
        .Top = Y
    End With
    
    Geometry_Create_Box Vertex2(), Outline, Outline, rgb_list2(), 0, 0
    Geometry_Create_Box vertex(), box_rect, box_rect, rgb_list(), 0, 0
    
    DirectDevice.SetTexture 0, Nothing
    DirectDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex2(0), Len(Vertex2(0))
    DirectDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, vertex(0), Len(vertex(0))
End Sub

Public Function ARGB(ByVal r As Long, ByVal g As Long, ByVal B As Long, ByVal A As Long) As Long
        
    Dim c As Long
        
    If A > 127 Then
        A = A - 128
        c = A * 2 ^ 24 Or &H80000000
        c = c Or r * 2 ^ 16
        c = c Or g * 2 ^ 8
        c = c Or B
    Else
        c = A * 2 ^ 24
        c = c Or r * 2 ^ 16
        c = c Or g * 2 ^ 8
        c = c Or B
    End If
    
    ARGB = c

End Function

Public Sub D3DColorToRgbList(rgb_list() As Long, color As D3DCOLORVALUE)
    rgb_list(0) = D3DColorARGB(color.A, color.r, color.g, color.B)
    rgb_list(1) = rgb_list(0)
    rgb_list(2) = rgb_list(0)
    rgb_list(3) = rgb_list(0)
End Sub

Public Sub Long_To_RGB_List(rgb_list() As Long, Long_Color As Long)
    rgb_list(0) = Long_Color
    rgb_list(1) = rgb_list(0)
    rgb_list(2) = rgb_list(0)
    rgb_list(3) = rgb_list(0)
End Sub

Private Function Char_Check(ByVal char_index As Integer) As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check char_index
    If char_index > 0 And char_index <= LastChar Then
        Char_Check = (charlist(char_index).Heading > 0)
    End If
    
End Function
 
Public Function Map_In_Bounds(ByVal map_x As Long, ByVal map_y As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Checks to see if a tile position is in the maps bounds
'*****************************************************************
    If map_x < map_current.map_x_min Or map_x > map_current.map_x_max Or map_y < map_current.map_y_min Or map_y > map_current.map_y_max Then
        Map_In_Bounds = False
        Exit Function
    End If
   
    Map_In_Bounds = True
End Function
Private Function LoadGrhData() As Boolean
On Error GoTo ErrorHandler
    Dim Grh As Long
    Dim Frame As Long
    Dim grhCount As Long
    Dim handle As Integer
    Dim fileVersion As Long
    Dim file As String
    
    file = Get_Extract(Scripts, "Graficos.ind")
    
    'Open files
    handle = FreeFile()
    Open file For Binary Access Read As handle
    Seek #handle, 1
   
    'Get file version
    Get handle, , fileVersion
   
    'Get number of grhs
    Get handle, , grhCount
   
    'Resize arrays
    ReDim GrhData(1 To grhCount) As GrhData
    
    Get handle, , Grh

    While Not Grh <= 0
        With GrhData(Grh)
        
            'Get number of frames
            Get handle, , .NumFrames
            If .NumFrames <= 0 Then GoTo ErrorHandler
            
           GrhData(Grh).Active = True

            ReDim .Frames(1 To GrhData(Grh).NumFrames)
           
            If .NumFrames > 1 Then
                'Read a animation GRH set
                For Frame = 1 To .NumFrames
                    Get handle, , .Frames(Frame)
                    If .Frames(Frame) <= 0 Or .Frames(Frame) > grhCount Then
                        GoTo ErrorHandler
                    End If
                Next Frame
               
                Get handle, , .Speed
               
                If .Speed <= 0 Then GoTo ErrorHandler
               
                'Compute width and height
                .pixelHeight = GrhData(.Frames(1)).pixelHeight
                If .pixelHeight <= 0 Then GoTo ErrorHandler
               
                .pixelWidth = GrhData(.Frames(1)).pixelWidth
                If .pixelWidth <= 0 Then GoTo ErrorHandler
               
                .TileWidth = GrhData(.Frames(1)).TileWidth
                If .TileWidth <= 0 Then GoTo ErrorHandler
               
                .TileHeight = GrhData(.Frames(1)).TileHeight
                If .TileHeight <= 0 Then GoTo ErrorHandler
            Else
                'Read in normal GRH data
                Get handle, , .FileNum
                If .FileNum <= 0 Then GoTo ErrorHandler
               
                Get handle, , GrhData(Grh).SX
                If .SX < 0 Then GoTo ErrorHandler
               
                Get handle, , .SY
                If .SY < 0 Then GoTo ErrorHandler
               
                Get handle, , .pixelWidth
                If .pixelWidth <= 0 Then GoTo ErrorHandler
               
                Get handle, , .pixelHeight
                If .pixelHeight <= 0 Then GoTo ErrorHandler
               
                'Compute width and height
                .TileWidth = .pixelWidth / 32
                .TileHeight = .pixelHeight / 32
               
                .Frames(1) = Grh
            End If
        End With
    Get handle, , Grh
    Wend
   
    Close handle
   Delete_File file
   
Dim count As Long
 
file = Get_Extract(Scripts, "minimap.dat")
Open file For Binary As #1
    Seek #1, 1
    For count = 1 To grhCount
        If GrhData(count).Active Then
            Get #1, , GrhData(count).MiniMap_color
        End If
    Next count
Close #1
Delete_File file

    LoadGrhData = True
Exit Function
 
ErrorHandler:
    LoadGrhData = False
End Function

'=-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-=
'=-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-=DIBUJA CUENTAS =-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-=
'=-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-=
'***********************************************************
Sub DibujarTodo(ByVal Index As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal casco As Integer, ByVal Shield As Integer, Weapon As Integer, ByVal Baned As Integer, ByVal nombre As String, ByVal LVL As Integer, ByVal Clase As String, ByVal Muerto As Integer, ByVal raza As String, ByVal Logged As Byte)

Dim Grh As Grh
Dim Pos As Integer
Dim YBody As Integer
Dim YYY As Integer
Dim XBody As Integer
Dim BBody As Integer

    'Mostramos la informacion de los personajes
    frmCuenta.nombre(Index).Caption = nombre
    
    If Logged = 0 Then
        frmCuenta.Label2(Index).Caption = ""
    Else
        frmCuenta.Label2(Index).Caption = "Logeado"
    End If
    
    frmCuenta.Label1(Index).Font = frmMain.Font
    frmCuenta.Label1(Index).Font = frmMain.Font
    
    frmCuenta.Label1(Index).Caption = LVL
    frmCuenta.Label2(Index).Caption = Clase
    frmCuenta.raza(Index).Caption = raza
    
    XBody = 12
    YBody = 20
    BBody = 17
    
    'Preparamos los datos para el fantasmita
    If Muerto = 1 Then
        Body = 8
        Head = 500
        Shield = 2
        Weapon = 2
        XBody = 10
        YBody = 35
        BBody = 16
    End If
    
    Grh = BodyData(Body).Walk(3)
    'Si no esta muerto lo mostramos como normalmente, pero si lo esta subimos el body
    Call DrawGrhtoHdc(frmCuenta.PJ(Index).hDC, BodyData(Body).Walk(3).GrhIndex, XBody, YBody, True)
    
    If Muerto = 0 Then YYY = BodyData(Body).HeadOffset.Y + 5
    If Muerto = 1 Then YYY = -9
    
    'Terminamos con el Body y vamos a por la cabeza
    Pos = YYY + GrhData(GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)).pixelHeight

    Call DrawGrhtoHdc(frmCuenta.PJ(Index).hDC, HeadData(Head).Head(3).GrhIndex, XBody + GrhData(BodyData(Body).Walk(3).GrhIndex).pixelWidth / 2 - GrhData(HeadData(Head).Head(3).GrhIndex).pixelWidth / 2, Pos, True)
            
    If casco <> 2 And casco > 0 Then
        Call DrawGrhtoHdc(frmCuenta.PJ(Index).hDC, CascoAnimData(casco).Head(3).GrhIndex, XBody - GrhData(CascoAnimData(casco).Head(3).GrhIndex).pixelWidth / 2 + GrhData(BodyData(Body).Walk(3).GrhIndex).pixelWidth / 2, Pos - GrhData(CascoAnimData(casco).Head(3).GrhIndex).pixelHeight + GrhData(HeadData(Head).Head(3).GrhIndex).pixelHeight, True)
    End If
    
    If Weapon <> 2 And Weapon > 0 Then
        Call DrawGrhtoHdc(frmCuenta.PJ(Index).hDC, WeaponAnimData(Weapon).WeaponWalk(3).GrhIndex, XBody, YBody, True)
    End If
                    
    If Shield <> 2 And Shield > 0 Then
        Call DrawGrhtoHdc(frmCuenta.PJ(Index).hDC, ShieldAnimData(Shield).ShieldWalk(3).GrhIndex, XBody, BBody, True)
    End If
        
    If Baned = 1 Then
       Call DrawGrhtoHdc(frmCuenta.PJ(Index).hDC, 20891, 0, 0, True)
    End If
        
End Sub
'***********************************************************
'=-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-=
'=-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-=/FIN DE DIBUJADO DE CUENTAS =-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-
'=-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-=

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''[PARTICULAS]''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Particle_Group_Create(ByVal map_x As Integer, ByVal map_y As Integer, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                        Optional ByVal particle_count As Long = 20, Optional ByVal stream_type As Long = 1, _
                                        Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                        Optional ByVal frame_speed As Single = 0.5, Optional ByVal ID As Long, _
                                        Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                                        Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                        Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                        Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                        Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                        Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                        Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                        Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                        Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                        Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                                        Optional grh_resizex As Integer, Optional grh_resizey As Integer, Optional ByVal Radio As Integer) As Long
                                        
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 12/15/2002
'Returns the particle_group_index if successful, else 0
'**************************************************************
    If (map_x <> -1) And (map_y <> -1) Then
    If Map_Particle_Group_Get(map_x, map_y) = 0 Then
        Particle_Group_Create = Particle_Group_Next_Open
        Particle_Group_Make Particle_Group_Create, map_x, map_y, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, ID, x1, y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey, Radio
    Else
        Particle_Group_Create = Particle_Group_Next_Open
        Particle_Group_Make Particle_Group_Create, map_x, map_y, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, ID, x1, y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey, Radio
    End If
    End If
End Function

Public Function Char_Particle_Group_Create(ByVal char_index As Integer, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                        Optional ByVal particle_count As Long = 20, Optional ByVal stream_type As Long = 1, _
                                        Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                        Optional ByVal frame_speed As Single = 0.5, Optional ByVal ID As Long, _
                                        Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                                        Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                        Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                        Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                        Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                        Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                        Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                        Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                        Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                        Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional Radio As Integer)
'**************************************************************
'Author: Augusto José Rando
'**************************************************************
    Dim char_part_free_index As Integer
    
    'If Char_Particle_Group_Find(char_index, stream_type) Then Exit Function ' hay que ver si dejar o sacar esto...
    If Not Char_Check(char_index) Then Exit Function
    char_part_free_index = Char_Particle_Group_Next_Open(char_index)
    
    If char_part_free_index > 0 Then
        Char_Particle_Group_Create = Particle_Group_Next_Open
        Char_Particle_Group_Make Char_Particle_Group_Create, char_index, char_part_free_index, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, ID, x1, y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, Radio
    End If

End Function
 
Public Function Particle_Group_Remove(ByVal particle_group_index As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    'Make sure it's a legal index
    If Particle_Group_Check(particle_group_index) Then
        Particle_Group_Destroy particle_group_index
        Particle_Group_Remove = True
    End If
End Function
 
Public Function Char_Particle_Group_Remove(ByVal char_index As Integer, ByVal stream_type As Long)
'**************************************************************
'Author: Augusto José Rando
'**************************************************************
    Dim char_part_index As Integer
    
    If Char_Check(char_index) Then
        char_part_index = Char_Particle_Group_Find(char_index, stream_type)
        If char_part_index = -1 Then Exit Function
        Call Particle_Group_Remove(char_part_index)
    End If

End Function
 
Public Function Particle_Group_Remove_All() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    Dim Index As Long
    
    For Index = 1 To particle_group_last
        'Make sure it's a legal index
        If Particle_Group_Check(Index) Then
            Particle_Group_Destroy Index
        End If
    Next Index
    
    Particle_Group_Remove_All = True
End Function

Public Function Char_Particle_Group_Remove_All(ByVal char_index As Integer)
'**************************************************************
'Author: Augusto José Rando
'**************************************************************
    Dim i As Integer
    
    If Char_Check(char_index) And Not charlist(char_index).particle_count = 0 Then
        For i = 1 To UBound(charlist(char_index).particle_group)
            If charlist(char_index).particle_group(i) <> 0 Then Call Particle_Group_Remove(charlist(char_index).particle_group(i))
        Next i
        Erase charlist(char_index).particle_group
        charlist(char_index).particle_count = 0
    End If
    
End Function
 
Public Function Particle_Group_Find(ByVal ID As Long) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'Find the index related to the handle
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    loopc = 1
    Do Until particle_group_list(loopc).ID = ID
        If loopc = particle_group_last Then
            Particle_Group_Find = 0
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Particle_Group_Find = loopc
Exit Function
ErrorHandler:
    Particle_Group_Find = 0
End Function
 
Private Function Char_Particle_Group_Find(ByVal char_index As Integer, ByVal stream_type As Long) As Integer
'*****************************************************************
'Author: Augusto José Rando
'Modified: returns slot or -1
'*****************************************************************

Dim i As Integer

For i = 1 To charlist(char_index).particle_count
    If particle_group_list(charlist(char_index).particle_group(i)).stream_type = stream_type Then
        Char_Particle_Group_Find = charlist(char_index).particle_group(i)
        Exit Function
    End If
Next i

Char_Particle_Group_Find = -1

End Function
Public Function Particle_Get_Type(ByVal particle_group_index As Long) As Byte
On Error GoTo ErrorHandler:
    Particle_Get_Type = particle_group_list(particle_group_index).stream_type
Exit Function
ErrorHandler:
    Particle_Get_Type = 0
End Function
Private Sub Particle_Group_Destroy(ByVal particle_group_index As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
On Error Resume Next
    Dim temp As particle_group
    Dim i As Integer
    
    If particle_group_list(particle_group_index).map_x > 0 And particle_group_list(particle_group_index).map_y > 0 Then
        MapData(particle_group_list(particle_group_index).map_x, particle_group_list(particle_group_index).map_y).particle_group_index = 0
    ElseIf particle_group_list(particle_group_index).char_index Then
        If Char_Check(particle_group_list(particle_group_index).char_index) Then
            For i = 1 To charlist(particle_group_list(particle_group_index).char_index).particle_count
                If charlist(particle_group_list(particle_group_index).char_index).particle_group(i) = particle_group_index Then
                    charlist(particle_group_list(particle_group_index).char_index).particle_group(i) = 0
                    Exit For
                End If
            Next i
        End If
    End If
    
    particle_group_list(particle_group_index) = temp
    
    'Update array size
    If particle_group_index = particle_group_last Then
        Do Until particle_group_list(particle_group_last).Active
            particle_group_last = particle_group_last - 1
            If particle_group_last = 0 Then
                particle_group_count = 0
                Exit Sub
            End If
        Loop
        Debug.Print particle_group_last & "," & UBound(particle_group_list)
        ReDim Preserve particle_group_list(1 To particle_group_last) As particle_group
    End If
    particle_group_count = particle_group_count - 1
End Sub

 
Private Sub Particle_Group_Make(ByVal particle_group_index As Long, ByVal map_x As Integer, ByVal map_y As Integer, _
                                ByVal particle_count As Long, ByVal stream_type As Long, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                Optional ByVal frame_speed As Single = 0.5, Optional ByVal ID As Long, _
                                Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                                Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                                Optional grh_resizex As Integer, Optional grh_resizey As Integer, Optional Radio As Integer)
                               
'*****************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Last Modify Date: 5/15/2003
'Makes a new particle effect
'Modified by Juan Martín Sotuyo Dodero
'*****************************************************************
    'Update array size
    If particle_group_index > particle_group_last Then
        particle_group_last = particle_group_index
        ReDim Preserve particle_group_list(1 To particle_group_last)
    End If
    particle_group_count = particle_group_count + 1
   
    'Make active
    particle_group_list(particle_group_index).Active = True
   
    'Map pos
    If (map_x <> -1) And (map_y <> -1) Then
        particle_group_list(particle_group_index).map_x = map_x
        particle_group_list(particle_group_index).map_y = map_y
    End If
   
    'Grh list
    ReDim particle_group_list(particle_group_index).grh_index_list(1 To UBound(grh_index_list))
    particle_group_list(particle_group_index).grh_index_list() = grh_index_list()
    particle_group_list(particle_group_index).grh_index_count = UBound(grh_index_list)
    
    particle_group_list(particle_group_index).Radio = Radio
   
    'Sets alive vars
    If alive_counter = -1 Then
        particle_group_list(particle_group_index).alive_counter = -1
        particle_group_list(particle_group_index).never_die = True
    Else
        particle_group_list(particle_group_index).alive_counter = alive_counter
        particle_group_list(particle_group_index).never_die = False
    End If
   
    'alpha blending
    particle_group_list(particle_group_index).alpha_blend = alpha_blend
   
    'stream type
    particle_group_list(particle_group_index).stream_type = stream_type
   
    'speed
    particle_group_list(particle_group_index).frame_speed = frame_speed
   
    particle_group_list(particle_group_index).x1 = x1
    particle_group_list(particle_group_index).y1 = y1
    particle_group_list(particle_group_index).x2 = x2
    particle_group_list(particle_group_index).y2 = y2
    particle_group_list(particle_group_index).angle = angle
    particle_group_list(particle_group_index).vecx1 = vecx1
    particle_group_list(particle_group_index).vecx2 = vecx2
    particle_group_list(particle_group_index).vecy1 = vecy1
    particle_group_list(particle_group_index).vecy2 = vecy2
    particle_group_list(particle_group_index).life1 = life1
    particle_group_list(particle_group_index).life2 = life2
    particle_group_list(particle_group_index).fric = fric
    particle_group_list(particle_group_index).spin = spin
    particle_group_list(particle_group_index).spin_speedL = spin_speedL
    particle_group_list(particle_group_index).spin_speedH = spin_speedH
    particle_group_list(particle_group_index).gravity = gravity
    particle_group_list(particle_group_index).grav_strength = grav_strength
    particle_group_list(particle_group_index).bounce_strength = bounce_strength
    particle_group_list(particle_group_index).XMove = XMove
    particle_group_list(particle_group_index).YMove = YMove
    particle_group_list(particle_group_index).move_x1 = move_x1
    particle_group_list(particle_group_index).move_x2 = move_x2
    particle_group_list(particle_group_index).move_y1 = move_y1
    particle_group_list(particle_group_index).move_y2 = move_y2
   
    particle_group_list(particle_group_index).rgb_list(0) = rgb_list(0)
    particle_group_list(particle_group_index).rgb_list(1) = rgb_list(1)
    particle_group_list(particle_group_index).rgb_list(2) = rgb_list(2)
    particle_group_list(particle_group_index).rgb_list(3) = rgb_list(3)
   
    'handle
    particle_group_list(particle_group_index).ID = ID
   
    'create particle stream
    particle_group_list(particle_group_index).particle_count = particle_count
    ReDim particle_group_list(particle_group_index).particle_stream(1 To particle_count)
   
    'plot particle group on map
    If (map_x <> -1) And (map_y <> -1) Then
        MapData(map_x, map_y).particle_group_index = particle_group_index
    End If
   
End Sub

Private Sub Char_Particle_Group_Make(ByVal particle_group_index As Long, ByVal char_index As Integer, ByVal particle_char_index As Integer, _
                                ByVal particle_count As Long, ByVal stream_type As Long, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                Optional ByVal frame_speed As Single = 0.5, Optional ByVal ID As Long, _
                                Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                                Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional Radio As Integer)
                                
'*****************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Last Modify Date: 5/15/2003
'Makes a new particle effect
'Modified by Juan Martín Sotuyo Dodero
'*****************************************************************
    'Update array size
    If particle_group_index > particle_group_last Then
        particle_group_last = particle_group_index
        ReDim Preserve particle_group_list(1 To particle_group_last)
    End If
    particle_group_count = particle_group_count + 1
    
    'Make active
    particle_group_list(particle_group_index).Active = True
    
    'Char index
    particle_group_list(particle_group_index).char_index = char_index
    
    'Grh list
    ReDim particle_group_list(particle_group_index).grh_index_list(1 To UBound(grh_index_list))
    particle_group_list(particle_group_index).grh_index_list() = grh_index_list()
    particle_group_list(particle_group_index).grh_index_count = UBound(grh_index_list)
    
    particle_group_list(particle_group_index).Radio = Radio
   
    'Sets alive vars
    If alive_counter = -1 Then
        particle_group_list(particle_group_index).alive_counter = -1
        particle_group_list(particle_group_index).never_die = True
    Else
        particle_group_list(particle_group_index).alive_counter = alive_counter
        particle_group_list(particle_group_index).never_die = False
    End If
   
    'alpha blending
    particle_group_list(particle_group_index).alpha_blend = alpha_blend
   
    'stream type
    particle_group_list(particle_group_index).stream_type = stream_type
   
    'speed
    particle_group_list(particle_group_index).frame_speed = frame_speed
   
    particle_group_list(particle_group_index).x1 = x1
    particle_group_list(particle_group_index).y1 = y1
    particle_group_list(particle_group_index).x2 = x2
    particle_group_list(particle_group_index).y2 = y2
    particle_group_list(particle_group_index).angle = angle
    particle_group_list(particle_group_index).vecx1 = vecx1
    particle_group_list(particle_group_index).vecx2 = vecx2
    particle_group_list(particle_group_index).vecy1 = vecy1
    particle_group_list(particle_group_index).vecy2 = vecy2
    particle_group_list(particle_group_index).life1 = life1
    particle_group_list(particle_group_index).life2 = life2
    particle_group_list(particle_group_index).fric = fric
    particle_group_list(particle_group_index).spin = spin
    particle_group_list(particle_group_index).spin_speedL = spin_speedL
    particle_group_list(particle_group_index).spin_speedH = spin_speedH
    particle_group_list(particle_group_index).gravity = gravity
    particle_group_list(particle_group_index).grav_strength = grav_strength
    particle_group_list(particle_group_index).bounce_strength = bounce_strength
    particle_group_list(particle_group_index).XMove = XMove
    particle_group_list(particle_group_index).YMove = YMove
    particle_group_list(particle_group_index).move_x1 = move_x1
    particle_group_list(particle_group_index).move_x2 = move_x2
    particle_group_list(particle_group_index).move_y1 = move_y1
    particle_group_list(particle_group_index).move_y2 = move_y2
   
    particle_group_list(particle_group_index).rgb_list(0) = rgb_list(0)
    particle_group_list(particle_group_index).rgb_list(1) = rgb_list(1)
    particle_group_list(particle_group_index).rgb_list(2) = rgb_list(2)
    particle_group_list(particle_group_index).rgb_list(3) = rgb_list(3)
   
    'handle
    particle_group_list(particle_group_index).ID = ID
   
    'create particle stream
    particle_group_list(particle_group_index).particle_count = particle_count
    ReDim particle_group_list(particle_group_index).particle_stream(1 To particle_count)
    
    'plot particle group on char
    charlist(char_index).particle_group(particle_char_index) = particle_group_index
   
End Sub

Public Function Particle_Type_Get(ByVal particle_index As Long) As Long
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last Modify Date: 8/27/2003
'Returns the stream type of a particle stream
'*****************************************************************
    If Particle_Group_Check(particle_index) Then
        Particle_Type_Get = particle_group_list(particle_index).stream_type
    Else
        Particle_Type_Get = 0
    End If
End Function
Private Sub Particle_Group_Render(ByVal particle_group_index As Long, ByVal screen_x As Long, ByVal screen_y As Long)
'*****************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Modified by: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/15/2003
'Renders a particle stream at a paticular screen point
'*****************************************************************

    Dim loopc As Long
    Dim temp_rgb(0 To 3) As Long
    Dim no_move As Boolean
    
    'Set colors
    If UserMinHP = 0 Then
        temp_rgb(0) = D3DColorARGB(particle_group_list(particle_group_index).alpha_blend, 255, 255, 255)
        temp_rgb(1) = D3DColorARGB(particle_group_list(particle_group_index).alpha_blend, 255, 255, 255)
        temp_rgb(2) = D3DColorARGB(particle_group_list(particle_group_index).alpha_blend, 255, 255, 255)
        temp_rgb(3) = D3DColorARGB(particle_group_list(particle_group_index).alpha_blend, 255, 255, 255)
    Else
        temp_rgb(0) = particle_group_list(particle_group_index).rgb_list(0)
        temp_rgb(1) = particle_group_list(particle_group_index).rgb_list(1)
        temp_rgb(2) = particle_group_list(particle_group_index).rgb_list(2)
        temp_rgb(3) = particle_group_list(particle_group_index).rgb_list(3)
    End If
    
    If particle_group_list(particle_group_index).alive_counter Then
    
        'See if it is time to move a particle
        particle_group_list(particle_group_index).frame_counter = particle_group_list(particle_group_index).frame_counter + timerTicksPerFrame
        If particle_group_list(particle_group_index).frame_counter > particle_group_list(particle_group_index).frame_speed Then
            particle_group_list(particle_group_index).frame_counter = 0
            no_move = False
        Else
            no_move = True
        End If
    
    
        'If it's still alive render all the particles inside
        For loopc = 1 To particle_group_list(particle_group_index).particle_count
                
        'Render particle
            Particle_Render particle_group_list(particle_group_index).particle_stream(loopc), _
                            screen_x, screen_y, _
                            particle_group_list(particle_group_index).grh_index_list(Round(RandomNumber(1, particle_group_list(particle_group_index).grh_index_count), 0)), _
                            temp_rgb(), _
                            particle_group_list(particle_group_index).alpha_blend, no_move, _
                            particle_group_list(particle_group_index).x1, particle_group_list(particle_group_index).y1, particle_group_list(particle_group_index).angle, _
                            particle_group_list(particle_group_index).vecx1, particle_group_list(particle_group_index).vecx2, _
                            particle_group_list(particle_group_index).vecy1, particle_group_list(particle_group_index).vecy2, _
                            particle_group_list(particle_group_index).life1, particle_group_list(particle_group_index).life2, _
                            particle_group_list(particle_group_index).fric, particle_group_list(particle_group_index).spin_speedL, _
                            particle_group_list(particle_group_index).gravity, particle_group_list(particle_group_index).grav_strength, _
                            particle_group_list(particle_group_index).bounce_strength, particle_group_list(particle_group_index).x2, _
                            particle_group_list(particle_group_index).y2, particle_group_list(particle_group_index).XMove, _
                            particle_group_list(particle_group_index).move_x1, particle_group_list(particle_group_index).move_x2, _
                            particle_group_list(particle_group_index).move_y1, particle_group_list(particle_group_index).move_y2, _
                            particle_group_list(particle_group_index).YMove, particle_group_list(particle_group_index).spin_speedH, _
                            particle_group_list(particle_group_index).spin, particle_group_list(particle_group_index).Radio, _
                            particle_group_list(particle_group_index).particle_count, loopc
        Next loopc
        
        If no_move = False Then
            'Update the group alive counter
            If particle_group_list(particle_group_index).never_die = False Then
                particle_group_list(particle_group_index).alive_counter = particle_group_list(particle_group_index).alive_counter - 1
            End If
        End If
    
    Else
        'If it's dead destroy it
        Particle_Group_Destroy particle_group_index
    End If
End Sub
 
Private Sub Particle_Render(ByRef temp_particle As Particle, ByVal screen_x As Long, ByVal screen_y As Long, _
                            ByVal grh_index As Long, ByRef rgb_list() As Long, _
                            Optional ByVal alpha_blend As Boolean, Optional ByVal no_move As Boolean, _
                            Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                            Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                            Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                            Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                            Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                            Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                            Optional ByVal bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                            Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                            Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                            Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, _
                            Optional ByVal Radio As Integer, Optional ByVal count As Integer, Optional ByVal Index As Integer)
'**************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Modified by: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/15/2003
'**************************************************************
    If no_move = False Then
        If temp_particle.alive_counter = 0 Then
            'Start new particle
            InitGrh temp_particle.Grh, grh_index, alpha_blend
            If Radio = 0 Then
                temp_particle.X = RandomNumber(x1, x2)
                temp_particle.Y = RandomNumber(y1, y2)
            Else
                temp_particle.X = (RandomNumber(x1, x2) + Radio) + Radio * Cos(PI * 2 * Index / count)
                temp_particle.Y = (RandomNumber(y1, y2) + Radio) + Radio * Sin(PI * 2 * Index / count)
            End If
            temp_particle.X = RandomNumber(x1, x2) - (base_tile_size \ 2)
            temp_particle.Y = RandomNumber(y1, y2) - (base_tile_size \ 2)
            temp_particle.vector_x = RandomNumber(vecx1, vecx2)
            temp_particle.vector_y = RandomNumber(vecy1, vecy2)
            temp_particle.angle = angle
            temp_particle.alive_counter = RandomNumber(life1, life2)
            temp_particle.friction = fric
        Else
            'Continue old particle
            'Do gravity
            If gravity = True Then
                temp_particle.vector_y = temp_particle.vector_y + grav_strength
                If temp_particle.Y > 0 Then
                    'bounce
                    temp_particle.vector_y = bounce_strength
                End If
            End If
            'Do rotation
            If spin = True Then temp_particle.Grh.angle = temp_particle.Grh.angle + (RandomNumber(spin_speedL, spin_speedH) / 100)
            If temp_particle.angle >= 360 Then
                temp_particle.angle = 0
            End If
            
            If XMove = True Then temp_particle.vector_x = RandomNumber(move_x1, move_x2)
            If YMove = True Then temp_particle.vector_y = RandomNumber(move_y1, move_y2)
        End If
        
        'Add in vector
        temp_particle.X = temp_particle.X + (temp_particle.vector_x \ temp_particle.friction)
        temp_particle.Y = temp_particle.Y + (temp_particle.vector_y \ temp_particle.friction)
    
        'decrement counter
         temp_particle.alive_counter = temp_particle.alive_counter - 1
    End If
    
'Draw it
    If temp_particle.Grh.GrhIndex Then
        DDrawTransGrhtoSurface temp_particle.Grh, temp_particle.X + screen_x, temp_particle.Y + screen_y, 1, 1, rgb_list(), 255, alpha_blend, temp_particle.Grh.angle
    End If
End Sub
Private Function Particle_Group_Next_Open() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    If particle_group_last = 0 Then
        Particle_Group_Next_Open = 1
        Exit Function
    End If
    
    loopc = 1
    Do Until particle_group_list(loopc).Active = False
        If loopc = particle_group_last Then
            Particle_Group_Next_Open = particle_group_last + 1
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Particle_Group_Next_Open = loopc

Exit Function

ErrorHandler:

End Function
 
Private Function Char_Particle_Group_Next_Open(ByVal char_index As Integer) As Integer
'*****************************************************************
'Author: Augusto José Rando
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    If charlist(char_index).particle_count = 0 Then
        Char_Particle_Group_Next_Open = charlist(char_index).particle_count + 1
        charlist(char_index).particle_count = Char_Particle_Group_Next_Open
        ReDim Preserve charlist(char_index).particle_group(1 To Char_Particle_Group_Next_Open) As Long
        Exit Function
    End If
    
    loopc = 1
    Do Until charlist(char_index).particle_group(loopc) = 0
        If loopc = charlist(char_index).particle_count Then
            Char_Particle_Group_Next_Open = charlist(char_index).particle_count + 1
            charlist(char_index).particle_count = Char_Particle_Group_Next_Open
            ReDim Preserve charlist(char_index).particle_group(1 To Char_Particle_Group_Next_Open) As Long
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Char_Particle_Group_Next_Open = loopc

Exit Function

ErrorHandler:
    charlist(char_index).particle_count = 1
    ReDim charlist(char_index).particle_group(1 To 1) As Long
    Char_Particle_Group_Next_Open = 1

End Function
 
Private Function Particle_Group_Check(ByVal particle_group_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check index
    If particle_group_index > 0 And particle_group_index <= particle_group_last Then
        If particle_group_list(particle_group_index).Active Then
            Particle_Group_Check = True
        End If
    End If
End Function

Public Function Map_Particle_Group_Get(ByVal map_x As Long, ByVal map_y As Long) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/20/2003
'Checks to see if a tile position has a particle_group_index and return it
'*****************************************************************
    If Map_In_Bounds(map_x, map_y) Then
        Map_Particle_Group_Get = map_current.map_grid(map_x, map_y).particle_group_index
    Else
        Map_Particle_Group_Get = 0
    End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''/[PARTICULAS]''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
