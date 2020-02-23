Attribute VB_Name = "modDirectDraw"
'**************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'**************************************************************

''
' modDirectDraw
'
' @remarks Funciones de DirectDraw y Visualizacion
' @author unkwown
' @version 0.0.20
' @date 20061015

Option Explicit

Function LoadWavetoDSBuffer(DS As DirectSound, DSB As DirectSoundBuffer, sFile As String) As Boolean
'*************************************************
'Author: Unkwown
'Last modified: 20/05/2006
'*************************************************
    
    Dim bufferDesc As DSBUFFERDESC
    Dim waveFormat As WAVEFORMATEX
    
    bufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
    
    waveFormat.nFormatTag = WAVE_FORMAT_PCM
    waveFormat.nChannels = 2
    waveFormat.lSamplesPerSec = 22050
    waveFormat.nBitsPerSample = 16
    waveFormat.nBlockAlign = waveFormat.nBitsPerSample / 8 * waveFormat.nChannels
    waveFormat.lAvgBytesPerSec = waveFormat.lSamplesPerSec * waveFormat.nBlockAlign
    Set DSB = DS.CreateSoundBufferFromFile(sFile, bufferDesc, waveFormat)
    
    If Err.Number <> 0 Then
        Exit Function
    End If
    
    LoadWavetoDSBuffer = True
    
End Function

Sub ConvertCPtoTP(StartPixelLeft As Integer, StartPixelTop As Integer, ByVal CX As Single, ByVal CY As Single, tX As Integer, tY As Integer)
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
Dim HWindowX As Integer
Dim HWindowY As Integer

CX = CX - StartPixelLeft
CY = CY - StartPixelTop

HWindowX = (WindowTileWidth \ 2)
HWindowY = (WindowTileHeight \ 2)

'Figure out X and Y tiles
CX = (CX \ TilePixelWidth)
CY = (CY \ TilePixelHeight)

If CX > HWindowX Then
    CX = (CX - HWindowX)

Else
    If CX < HWindowX Then
        CX = (0 - (HWindowX - CX))
    Else
        CX = 0
    End If
End If

If CY > HWindowY Then
    CY = (0 - (HWindowY - CY))
Else
    If CY < HWindowY Then
        CY = (CY - HWindowY)
    Else
        CY = 0
    End If
End If

tX = UserPos.X + CX
tY = UserPos.y + CY

End Sub




Function DeInitTileEngine() As Boolean
'*************************************************
'Author: Unkwown
'Last modified: 26/05/06
'*************************************************
Dim loopc As Integer

EngineRun = False

'****** Clear DirectX objects ******
Set PrimarySurface = Nothing
Set PrimaryClipper = Nothing
Set BackBufferSurface = Nothing

Set SurfaceDB = Nothing

Set DirectDraw = Nothing

'Reset any channels that are done
For loopc = 1 To NumSoundBuffers
    Set DSBuffers(loopc) = Nothing
Next loopc

Set DirectSound = Nothing

Set DirectX = Nothing

DeInitTileEngine = True

End Function


Sub MakeChar(CharIndex As Integer, Body As Integer, Head As Integer, Heading As Byte, X As Integer, y As Integer)
'*************************************************
'Author: Unkwown
'Last modified: 28/05/06 by GS
'*************************************************
On Error Resume Next

'Update LastChar
If CharIndex > LastChar Then LastChar = CharIndex
NumChars = NumChars + 1

'Update head, body, ect.
CharList(CharIndex).Body = BodyData(Body)
CharList(CharIndex).Head = HeadData(Head)
CharList(CharIndex).Heading = Heading

'Reset moving stats
CharList(CharIndex).Moving = 0
CharList(CharIndex).MoveOffset.X = 0
CharList(CharIndex).MoveOffset.y = 0

'Update position
CharList(CharIndex).Pos.X = X
CharList(CharIndex).Pos.y = y

'Make active
CharList(CharIndex).Active = 1

'Plot on map
MapData(X, y).CharIndex = CharIndex

bRefreshRadar = True ' GS

End Sub







Sub EraseChar(CharIndex As Integer)
'*************************************************
'Author: Unkwown
'Last modified: 28/05/06 by GS
'*************************************************
If CharIndex = 0 Then Exit Sub
'Make un-active
CharList(CharIndex).Active = 0

'Update lastchar
If CharIndex = LastChar Then
    Do Until CharList(LastChar).Active = 1
        LastChar = LastChar - 1
        If LastChar = 0 Then Exit Do
    Loop
End If

MapData(CharList(CharIndex).Pos.X, CharList(CharIndex).Pos.y).CharIndex = 0

'Update NumChars
NumChars = NumChars - 1

bRefreshRadar = True ' GS

End Sub

Sub InitGrh(ByRef Grh As Grh, ByVal GrhIndex As Integer, Optional Started As Byte = 2)
'*************************************************
'Author: Unkwown
'Last modified: 31/05/06 - GS
'*************************************************
On Error Resume Next
Grh.GrhIndex = GrhIndex
If Grh.GrhIndex <> 0 Then ' 31/05/2006
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
    Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
Else
    Grh.FrameCounter = 1
    Grh.Started = 0
    Grh.SpeedCounter = 0
End If

End Sub

Sub MoveCharbyHead(CharIndex As Integer, nHeading As Byte)
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
Dim addX As Integer
Dim addY As Integer
Dim X As Integer
Dim y As Integer
Dim nX As Integer
Dim nY As Integer

X = CharList(CharIndex).Pos.X
y = CharList(CharIndex).Pos.y

'Figure out which way to move
Select Case nHeading

    Case NORTH
        addY = -1

    Case EAST
        addX = 1

    Case SOUTH
        addY = 1
    
    Case WEST
        addX = -1
        
End Select

nX = X + addX
nY = y + addY

MapData(nX, nY).CharIndex = CharIndex
CharList(CharIndex).Pos.X = nX
CharList(CharIndex).Pos.y = nY
MapData(X, y).CharIndex = 0

CharList(CharIndex).MoveOffset.X = -1 * (TilePixelWidth * addX)
CharList(CharIndex).MoveOffset.y = -1 * (TilePixelHeight * addY)

CharList(CharIndex).Moving = 1
CharList(CharIndex).Heading = nHeading

End Sub

Sub MoveCharbyPos(CharIndex As Integer, nX As Integer, nY As Integer)
'*************************************************
'Author: Unkwown
'Last modified: 28/05/06 by GS
'*************************************************
Dim X As Integer
Dim y As Integer
Dim addX As Integer
Dim addY As Integer
Dim nHeading As Byte

X = CharList(CharIndex).Pos.X
y = CharList(CharIndex).Pos.y

addX = nX - X
addY = nY - y

If Sgn(addX) = 1 Then
    nHeading = EAST
End If

If Sgn(addX) = -1 Then
    nHeading = WEST
End If

If Sgn(addY) = -1 Then
    nHeading = NORTH
End If

If Sgn(addY) = 1 Then
    nHeading = SOUTH
End If

MapData(nX, nY).CharIndex = CharIndex
CharList(CharIndex).Pos.X = nX
CharList(CharIndex).Pos.y = nY
MapData(X, y).CharIndex = 0

CharList(CharIndex).MoveOffset.X = -1 * (TilePixelWidth * addX)
CharList(CharIndex).MoveOffset.y = -1 * (TilePixelHeight * addY)

CharList(CharIndex).Moving = 1
CharList(CharIndex).Heading = nHeading

bRefreshRadar = True ' GS

End Sub


Function NextOpenChar() As Integer
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
Dim loopc As Integer

loopc = 1
Do While CharList(loopc).Active
    loopc = loopc + 1
Loop

NextOpenChar = loopc

End Function

Function LegalPos(X As Integer, y As Integer) As Boolean
'*************************************************
'Author: Unkwown
'Last modified: 28/05/06 - GS
'*************************************************

LegalPos = True

'Check to see if its out of bounds
If X - 8 < 1 Or X - 8 > 100 Or y - 6 < 1 Or y - 6 > 100 Then
    LegalPos = False
    Exit Function
End If

'Check to see if its blocked
If MapData(X, y).Blocked = 1 Then
    LegalPos = False
    Exit Function
End If

'Check for character
If MapData(X, y).CharIndex > 0 Then
    LegalPos = False
    Exit Function
End If

End Function




Function InMapLegalBounds(X As Integer, y As Integer) As Boolean
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************

If X < MinXBorder Or X > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
    InMapLegalBounds = False
    Exit Function
End If

InMapLegalBounds = True

End Function

Function InMapBounds(X As Integer, y As Integer) As Boolean
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************

If X < XMinMapSize Or X > XMaxMapSize Or y < YMinMapSize Or y > YMaxMapSize Then
    InMapBounds = False
    Exit Function
End If

InMapBounds = True

End Function

Sub DDrawTransGrhtoSurface(ByRef Surface As DirectDrawSurface7, Grh As Grh, ByVal X As Integer, ByVal y As Integer, Center As Byte, Animate As Byte, Optional ByVal KillAnim As Integer = 0)
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
If MapaCargado = False Then Exit Sub

Dim iGrhIndex As Integer
Dim SourceRect As RECT
Dim QuitarAnimacion As Boolean

If Grh.GrhIndex = 0 Then Exit Sub

'Figure out what frame to draw (always 1 if not animated)
iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
If iGrhIndex = 0 Then Exit Sub
'Center Grh over X,Y pos
If Center Then
    If GrhData(iGrhIndex).TileWidth <> 1 Then
        X = X - Int(GrhData(iGrhIndex).TileWidth * 16) + 16 'hard coded for speed
    End If
    If GrhData(iGrhIndex).TileHeight <> 1 Then
        y = y - Int(GrhData(iGrhIndex).TileHeight * 32) + 32 'hard coded for speed
    End If
End If

With SourceRect
    .Left = GrhData(iGrhIndex).sX
    .Top = GrhData(iGrhIndex).sY
    .Right = .Left + GrhData(iGrhIndex).pixelWidth
    .Bottom = .Top + GrhData(iGrhIndex).pixelHeight
End With '
Surface.BltFast X, y, SurfaceDB.Surface(GrhData(iGrhIndex).FileNum), SourceRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY

End Sub
Sub DibujarGrhInidex(ByRef Surface As DirectDrawSurface7, iGrhIndex As Integer, ByVal X As Integer, ByVal y As Integer, Center As Byte, Animate As Byte, Optional ByVal KillAnim As Integer = 0, Optional Alpha As Byte = 200)
'*************************************************
'Author: Loopzer
'Last modified: 20/11/07
'*************************************************
If MapaCargado = False Then Exit Sub


Dim SourceRect As RECT
Dim QuitarAnimacion As Boolean



'Figure out what frame to draw (always 1 if not animated)

If iGrhIndex = 0 Then Exit Sub
'Center Grh over X,Y pos
If Center Then
    If GrhData(iGrhIndex).TileWidth <> 1 Then
        X = X - Int(GrhData(iGrhIndex).TileWidth * 16) + 16 'hard coded for speed
    End If
    If GrhData(iGrhIndex).TileHeight <> 1 Then
        y = y - Int(GrhData(iGrhIndex).TileHeight * 32) + 32 'hard coded for speed
    End If
End If

With SourceRect
    .Left = GrhData(iGrhIndex).sX
    .Top = GrhData(iGrhIndex).sY
    .Right = .Left + GrhData(iGrhIndex).pixelWidth
    .Bottom = .Top + GrhData(iGrhIndex).pixelHeight
End With '
'MotorDeEfectos.DBAlpha SurfaceDB.Surface(GrhData(iGrhIndex).FileNum), SourceRect, False, X, Y, Alpha
Surface.BltFast X, y, SurfaceDB.Surface(GrhData(iGrhIndex).FileNum), SourceRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY

End Sub




Sub DrawBackBufferSurface()
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
PrimarySurface.Blt MainViewRect, BackBufferSurface, MainDestRect, DDBLT_WAIT
End Sub




Sub DrawGrhtoHdc(hWnd As Long, hdc As Long, Grh As Integer, SourceRect As RECT, destRect As RECT)
'*************************************************
'Author: Unkwown
'Last modified: 26/05/06 - GS
'*************************************************
On Error Resume Next
If Grh <= 0 Then Exit Sub
Dim aux As Integer
aux = GrhData(Grh).FileNum
If aux = 0 Then Exit Sub
SecundaryClipper.SetHWnd hWnd
SurfaceDB.Surface(aux).BltToDC hdc, SourceRect, destRect
End Sub

Sub PlayWaveDS(file As String)
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************

    'Cylce through avaiable sound buffers
    LastSoundBufferUsed = LastSoundBufferUsed + 1
    If LastSoundBufferUsed > NumSoundBuffers Then
        LastSoundBufferUsed = 1
    End If
    
    If LoadWavetoDSBuffer(DirectSound, DSBuffers(LastSoundBufferUsed), file) Then
        DSBuffers(LastSoundBufferUsed).Play DSBPLAY_DEFAULT
    End If

End Sub
' [Loopzer]
Public Sub DePegar()
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
    Dim X As Integer
    Dim y As Integer

    For X = 0 To DeSeleccionAncho - 1
        For y = 0 To DeSeleccionAlto - 1
             MapData(X + DeSeleccionOX, y + DeSeleccionOY) = DeSeleccionMap(X, y)
        Next
    Next
End Sub
Public Sub PegarSeleccion() '(mx As Integer, my As Integer)
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
    'podria usar copy mem , pero por las dudas no XD
    Static UltimoX As Integer
    Static UltimoY As Integer
    If UltimoX = SobreX And UltimoY = SobreY Then Exit Sub
    UltimoX = SobreX
    UltimoY = SobreY
    Dim X As Integer
    Dim y As Integer
    DeSeleccionAncho = SeleccionAncho
    DeSeleccionAlto = SeleccionAlto
    DeSeleccionOX = SobreX
    DeSeleccionOY = SobreY
    ReDim DeSeleccionMap(DeSeleccionAncho, DeSeleccionAlto) As MapBlock
    
    For X = 0 To DeSeleccionAncho - 1
        For y = 0 To DeSeleccionAlto - 1
            DeSeleccionMap(X, y) = MapData(X + SobreX, y + SobreY)
        Next
    Next
    For X = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1
             MapData(X + SobreX, y + SobreY) = SeleccionMap(X, y)
        Next
    Next
    Seleccionando = False
End Sub
Public Sub AccionSeleccion()
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
    Dim X As Integer
    Dim y As Integer
    SeleccionAncho = Abs(SeleccionIX - SeleccionFX) + 1
    SeleccionAlto = Abs(SeleccionIY - SeleccionFY) + 1
    DeSeleccionAncho = SeleccionAncho
    DeSeleccionAlto = SeleccionAlto
    DeSeleccionOX = SeleccionIX
    DeSeleccionOY = SeleccionIY
    ReDim DeSeleccionMap(DeSeleccionAncho, DeSeleccionAlto) As MapBlock
    
    For X = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1
            DeSeleccionMap(X, y) = MapData(X + SeleccionIX, y + SeleccionIY)
        Next
    Next
    For X = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1
           ClickEdit vbLeftButton, SeleccionIX + X, SeleccionIY + y
        Next
    Next
    Seleccionando = False
End Sub

Public Sub BlockearSeleccion()
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
    Dim X As Integer
    Dim y As Integer
    Dim Vacio As MapBlock
    SeleccionAncho = Abs(SeleccionIX - SeleccionFX) + 1
    SeleccionAlto = Abs(SeleccionIY - SeleccionFY) + 1
    DeSeleccionAncho = SeleccionAncho
    DeSeleccionAlto = SeleccionAlto
    DeSeleccionOX = SeleccionIX
    DeSeleccionOY = SeleccionIY
    ReDim DeSeleccionMap(DeSeleccionAncho, DeSeleccionAlto) As MapBlock
    
    For X = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1
            DeSeleccionMap(X, y) = MapData(X + SeleccionIX, y + SeleccionIY)
        Next
    Next
    For X = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1
             If MapData(X + SeleccionIX, y + SeleccionIY).Blocked = 1 Then
                MapData(X + SeleccionIX, y + SeleccionIY).Blocked = 0
             Else
                MapData(X + SeleccionIX, y + SeleccionIY).Blocked = 1
            End If
        Next
    Next
    Seleccionando = False
End Sub
Public Sub CortarSeleccion()
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
    CopiarSeleccion
    Dim X As Integer
    Dim y As Integer
    Dim Vacio As MapBlock
    DeSeleccionAncho = SeleccionAncho
    DeSeleccionAlto = SeleccionAlto
    DeSeleccionOX = SeleccionIX
    DeSeleccionOY = SeleccionIY
    ReDim DeSeleccionMap(DeSeleccionAncho, DeSeleccionAlto) As MapBlock
    
    For X = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1
            DeSeleccionMap(X, y) = MapData(X + SeleccionIX, y + SeleccionIY)
        Next
    Next
    For X = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1
             MapData(X + SeleccionIX, y + SeleccionIY) = Vacio
        Next
    Next
    Seleccionando = False
End Sub
Public Sub CopiarSeleccion()
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
    'podria usar copy mem , pero por las dudas no XD
    Dim X As Integer
    Dim y As Integer
    Seleccionando = False
    SeleccionAncho = Abs(SeleccionIX - SeleccionFX) + 1
    SeleccionAlto = Abs(SeleccionIY - SeleccionFY) + 1
    ReDim SeleccionMap(SeleccionAncho, SeleccionAlto) As MapBlock
    For X = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1
            SeleccionMap(X, y) = MapData(X + SeleccionIX, y + SeleccionIY)
        Next
    Next
End Sub
Public Sub GenerarVista()
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
   ' hacer una llamada a un seter o geter , es mas lento q una variable
   ' con esto hacemos q no este preguntando a el objeto cadavez
   ' q dibuja , Render mas rapido ;)
    VerBlockeados = frmMain.cVerBloqueos.value
    VerTriggers = frmMain.cVerTriggers.value
    VerCapa1 = frmMain.mnuVerCapa1.Checked
    VerCapa2 = frmMain.mnuVerCapa2.Checked
    VerCapa3 = frmMain.mnuVerCapa3.Checked
    VerCapa4 = frmMain.mnuVerCapa4.Checked
    VerTranslados = frmMain.mnuVerTranslados.Checked
    VerObjetos = frmMain.mnuVerObjetos.Checked
    VerNpcs = frmMain.mnuVerNPCs.Checked
    
End Sub
' [/Loopzer]
Public Sub RenderScreen(TileX As Integer, TileY As Integer, PixelOffsetX As Integer, PixelOffsetY As Integer)
'*************************************************
'Author: Unkwown
'Last modified: 31/05/06 by GS
'Last modified: 21/11/07 By Loopzer
'Last modifier: 24/11/08 by GS
'*************************************************

On Error Resume Next
Dim y       As Integer              'Keeps track of where on map we are
Dim X       As Integer
Dim minY    As Integer              'Start Y pos on current map
Dim maxY    As Integer              'End Y pos on current map
Dim minX    As Integer              'Start X pos on current map
Dim maxX    As Integer              'End X pos on current map
Dim ScreenX As Integer              'Keeps track of where to place tile on screen
Dim ScreenY As Integer
Dim r       As RECT
Dim Sobre   As Integer
Dim Moved   As Byte
Dim iPPx    As Integer              'Usado en el Layer de Chars
Dim iPPy    As Integer              'Usado en el Layer de Chars
Dim Grh     As Grh                  'Temp Grh for show tile and blocked
Dim bCapa    As Byte                 'cCapas ' 31/05/2006 - GS, control de Capas
Dim SelRect As RECT
Dim rSourceRect         As RECT     'Usado en el Layer 1
Dim iGrhIndex           As Integer  'Usado en el Layer 1
Dim PixelOffsetXTemp    As Integer  'For centering grhs
Dim PixelOffsetYTemp    As Integer
Dim TempChar            As Char
BackBufferSurface.BltColorFill r, 0 'Solucion a algunos temas molestos :P
minY = (TileY - (WindowTileHeight \ 2)) - TileBufferSize
maxY = (TileY + (WindowTileHeight \ 2)) + TileBufferSize
minX = (TileX - (WindowTileWidth \ 2)) - TileBufferSize
maxX = (TileX + (WindowTileWidth \ 2)) + TileBufferSize
' 31/05/2006 - GS, control de Capas
If Val(frmMain.cCapas.Text) >= 1 And (frmMain.cCapas.Text) <= 4 Then
    bCapa = Val(frmMain.cCapas.Text)
Else
    bCapa = 1
End If
GenerarVista 'Loopzer
ScreenY = 8
For y = (minY + 8) To (maxY - 8)
    ScreenX = 8
    For X = (minX + 8) To (maxX - 8)
        If InMapBounds(X, y) Then
            If X > 100 Or y < 1 Then Exit For ' 30/05/2006

            'Layer 1 **********************************
            If SobreX = X And SobreY = y Then
                ' Pone Grh !
                Sobre = -1
                If frmMain.cSeleccionarSuperficie.value = True Then
                    Sobre = MapData(X, y).Graphic(bCapa).GrhIndex
                    If frmConfigSup.MOSAICO.value = vbChecked Then
                        Dim aux As Integer
                        Dim dy As Integer
                        Dim dX As Integer
                        If frmConfigSup.DespMosaic.value = vbChecked Then
                            dy = Val(frmConfigSup.DMLargo.Text)
                            dX = Val(frmConfigSup.DMAncho.Text)
                        Else
                            dy = 0
                            dX = 0
                        End If
                        If frmMain.mnuAutoCompletarSuperficies.Checked = False Then
                            aux = Val(frmMain.cGrh.Text) + _
                            (((y + dy) Mod frmConfigSup.mLargo.Text) * frmConfigSup.mAncho.Text) + ((X + dX) Mod frmConfigSup.mAncho.Text)
                            If MapData(X, y).Graphic(bCapa).GrhIndex <> aux Then
                                MapData(X, y).Graphic(bCapa).GrhIndex = aux
                                InitGrh MapData(X, y).Graphic(bCapa), aux
                            End If
                        Else
                            aux = Val(frmMain.cGrh.Text) + _
                            (((y + dy) Mod frmConfigSup.mLargo.Text) * frmConfigSup.mAncho.Text) + ((X + dX) Mod frmConfigSup.mAncho.Text)
                            If MapData(X, y).Graphic(bCapa).GrhIndex <> aux Then
                                MapData(X, y).Graphic(bCapa).GrhIndex = aux
                                InitGrh MapData(X, y).Graphic(bCapa), aux
                            End If
                        End If
                    Else
                        If MapData(X, y).Graphic(bCapa).GrhIndex <> Val(frmMain.cGrh.Text) Then
                            MapData(X, y).Graphic(bCapa).GrhIndex = Val(frmMain.cGrh.Text)
                            InitGrh MapData(X, y).Graphic(bCapa), Val(frmMain.cGrh.Text)
                        End If
                    End If
                End If
            Else
                Sobre = -1
            End If
            If VerCapa1 Then
            With MapData(X, y).Graphic(1)
                If (.GrhIndex <> 0) Then
                    If (.Started = 1) Then
                        If (.SpeedCounter > 0) Then
                            .SpeedCounter = .SpeedCounter - 1
                            If (.SpeedCounter = 0) Then
                                .SpeedCounter = GrhData(.GrhIndex).Speed
                                .FrameCounter = .FrameCounter + 1
                                If (.FrameCounter > GrhData(.GrhIndex).NumFrames) Then _
                                    .FrameCounter = 1
                            End If
                        End If
                    End If
                    'Figure out what frame to draw (always 1 if not animated)
                    iGrhIndex = GrhData(.GrhIndex).Frames(.FrameCounter)
                End If
            End With
            If iGrhIndex <> 0 Then
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
            End If
            End If
            'Layer 2 **********************************
            If MapData(X, y).Graphic(2).GrhIndex <> 0 And VerCapa2 Then
                Call DDrawTransGrhtoSurface( _
                        BackBufferSurface, _
                        MapData(X, y).Graphic(2), _
                        ((32 * ScreenX) - 32) + PixelOffsetX, _
                        ((32 * ScreenY) - 32) + PixelOffsetY, _
                        1, _
                        1)
            End If
            If Sobre >= 0 Then
                If MapData(X, y).Graphic(bCapa).GrhIndex <> Sobre Then
                MapData(X, y).Graphic(bCapa).GrhIndex = Sobre
                InitGrh MapData(X, y).Graphic(bCapa), Sobre
                End If
            End If
        End If
        ScreenX = ScreenX + 1
    Next X
    ScreenY = ScreenY + 1
    If y > 100 Then Exit For
Next y
ScreenY = 8
For y = (minY + 8) To (maxY - 1) '- 8+ 8
    ScreenX = 5
    For X = (minX + 5) To (maxX - 5) '- 8 + 8
        If InMapBounds(X, y) Then
            If X > 100 Or X < -3 Then Exit For ' 30/05/2006

            iPPx = ((32 * ScreenX) - 32) + PixelOffsetX
            iPPy = ((32 * ScreenY) - 32) + PixelOffsetY
             'Object Layer **********************************
             If MapData(X, y).OBJInfo.objindex <> 0 And VerObjetos Then
                 Call DDrawTransGrhtoSurface( _
                         BackBufferSurface, _
                         MapData(X, y).ObjGrh, _
                         iPPx, iPPy, 1, 1)
             End If
            
                  'Char layer **********************************
                 If MapData(X, y).CharIndex <> 0 And VerNpcs Then
                 
                     TempChar = CharList(MapData(X, y).CharIndex)
                 
                     PixelOffsetXTemp = PixelOffsetX
                     PixelOffsetYTemp = PixelOffsetY
                    
                   'Dibuja solamente players
                   If TempChar.Head.Head(TempChar.Heading).GrhIndex <> 0 Then
                     'Draw Body
                     Call DDrawTransGrhtoSurface(BackBufferSurface, TempChar.Body.Walk(TempChar.Heading), (PixelPos(ScreenX) + PixelOffsetXTemp), PixelPos(ScreenY) + PixelOffsetYTemp, 1, 1)
                     'Draw Head
                     Call DDrawTransGrhtoSurface(BackBufferSurface, TempChar.Head.Head(TempChar.Heading), (PixelPos(ScreenX) + PixelOffsetXTemp) + TempChar.Body.HeadOffset.X, PixelPos(ScreenY) + PixelOffsetYTemp + TempChar.Body.HeadOffset.y, 1, 0)
                   Else: Call DDrawTransGrhtoSurface(BackBufferSurface, TempChar.Body.Walk(TempChar.Heading), (PixelPos(ScreenX) + PixelOffsetXTemp), PixelPos(ScreenY) + PixelOffsetYTemp, 1, 1)
                   End If
                 End If
             'Layer 3 *****************************************
             If MapData(X, y).Graphic(3).GrhIndex <> 0 And VerCapa3 Then
                 'Draw
                 'Call DDrawTransGrhtoSurface( _
                         BackBufferSurface, _
                         MapData(X, Y).Graphic(3), _
                         ((32 * ScreenX) - 32) + PixelOffsetX, _
                         ((32 * ScreenY) - 32) + PixelOffsetY, _
                         1, 1)
                         Call DDrawTransGrhtoSurface( _
                         BackBufferSurface, _
                         MapData(X, y).Graphic(3), _
                         iPPx, _
                         iPPy, _
                         1, 1)
             End If
        End If
        ScreenX = ScreenX + 1
    Next X
    ScreenY = ScreenY + 1
Next y
'Tiles blokeadas, techos, triggers , seleccion
ScreenY = 5
For y = (minY + 5) To (maxY - 1)
    ScreenX = 5
    For X = (minX + 5) To (maxX)
        If X < 101 And X > 0 And y < 101 And y > 0 Then ' 30/05/2006
            iPPx = ((32 * ScreenX) - 32) + PixelOffsetX
            iPPy = ((32 * ScreenY) - 32) + PixelOffsetY
            If MapData(X, y).Graphic(4).GrhIndex <> 0 _
            And (frmMain.mnuVerCapa4.Checked = True) Then
                'Draw
                Call DDrawTransGrhtoSurface( _
                    BackBufferSurface, _
                    MapData(X, y).Graphic(4), _
                    iPPx, _
                    iPPy, _
                    1, 1)
            End If
            If MapData(X, y).TileExit.Map <> 0 And VerTranslados Then
                Grh.GrhIndex = 3
                Grh.FrameCounter = 1
                Grh.Started = 0
                Call DDrawTransGrhtoSurface( _
                    BackBufferSurface, _
                    Grh, _
                    iPPx, _
                    iPPy, _
                    1, 1)
            End If
            'Show blocked tiles
            If VerBlockeados And MapData(X, y).Blocked = 1 Then
                'Grh.GrhIndex = 4
                'Grh.FrameCounter = 1
                'Grh.Started = 0
                'Call DDrawTransGrhtoSurface( _
                '    BackBufferSurface, _
                '    Grh, _
                '    ((32 * ScreenX) - 32) + PixelOffsetX, _
                '    ((32 * ScreenY) - 32) + PixelOffsetY, _
                '    1, 1)
                BackBufferSurface.SetForeColor vbWhite
                BackBufferSurface.SetFillColor vbRed
                BackBufferSurface.SetFillStyle 0
                Call BackBufferSurface.DrawBox( _
                    iPPx + 16, _
                    iPPy + 16, _
                    iPPx + 21, _
                    iPPy + 21)
            End If
            If VerGrilla Then
                ' Grilla 24/11/2008 by GS
                BackBufferSurface.SetForeColor vbRed
                BackBufferSurface.DrawLine ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, iPPx, iPPy + 32
                BackBufferSurface.DrawLine ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, iPPx + 32, iPPy
            End If
            If VerTriggers Then
                Call DrawText(PixelPos(ScreenX), PixelPos(ScreenY), Str(MapData(X, y).Trigger), vbRed)
            End If
            If Seleccionando Then
                'If ScreenX >= SeleccionIX And ScreenX <= SeleccionFX And ScreenY >= SeleccionIY And ScreenY <= SeleccionFY Then
                    If X >= SeleccionIX And y >= SeleccionIY Then
                        If X <= SeleccionFX And y <= SeleccionFY Then
                            BackBufferSurface.SetForeColor vbGreen
                            BackBufferSurface.SetFillStyle 1
                            BackBufferSurface.DrawBox iPPx, iPPy, iPPx + 32, iPPy + 32
                        End If
                    End If
            End If

        End If
        ScreenX = ScreenX + 1
    Next X
    ScreenY = ScreenY + 1
Next y

End Sub



Public Sub DrawText(lngXPos As Integer, lngYPos As Integer, strText As String, lngColor As Long)
'*************************************************
'Author: Unkwown
'Last modified: 26/05/06
'*************************************************
   If LenB(strText) <> 0 Then
    'BackBufferSurface.SetFontTransparency True                           'Set the transparency flag to true
    'BackBufferSurface.SetForeColor vbBlack                               'Set the color of the text to the color passed to the sub
    'BackBufferSurface.SetFont frmMain.Font                               'Set the font used to the font on the form
    'BackBufferSurface.DrawText lngXPos - 2, lngYPos - 1, strText, False  'Draw the text on to the screen, in the coordinates specified
    
    
    BackBufferSurface.SetFontTransparency True                           'Set the transparency flag to true
    BackBufferSurface.SetForeColor lngColor                              'Set the color of the text to the color passed to the sub
    BackBufferSurface.SetFont frmMain.Font                               'Set the font used to the font on the form
    BackBufferSurface.DrawText lngXPos, lngYPos, strText, False          'Draw the text on to the screen, in the coordinates specified
   End If
End Sub

Function HayUserAbajo(X As Integer, y As Integer, GrhIndex) As Boolean
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
HayUserAbajo = _
    CharList(UserCharIndex).Pos.X >= X - (GrhData(GrhIndex).TileWidth \ 2) _
And CharList(UserCharIndex).Pos.X <= X + (GrhData(GrhIndex).TileWidth \ 2) _
And CharList(UserCharIndex).Pos.y >= y - (GrhData(GrhIndex).TileHeight - 1) _
And CharList(UserCharIndex).Pos.y <= y
End Function



Function PixelPos(X As Integer) As Integer
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************

PixelPos = (TilePixelWidth * X) - TilePixelWidth

End Function

Function InitTileEngine(ByRef setDisplayFormhWnd As Long, setMainViewTop As Integer, setMainViewLeft As Integer, setTilePixelHeight As Integer, setTilePixelWidth As Integer, setWindowTileHeight As Integer, setWindowTileWidth As Integer, setTileBufferSize As Integer) As Boolean
'*************************************************
'Author: Unkwown
'Last modified: 15/10/06 by GS
'*************************************************

Dim SurfaceDesc As DDSURFACEDESC2
Dim ddck As DDCOLORKEY

'Fill startup variables
DisplayFormhWnd = setDisplayFormhWnd
MainViewTop = setMainViewTop
MainViewLeft = setMainViewLeft
TilePixelWidth = setTilePixelWidth
TilePixelHeight = setTilePixelHeight
WindowTileHeight = setWindowTileHeight
WindowTileWidth = setWindowTileWidth
TileBufferSize = setTileBufferSize

'[GS] 02/10/2006
MinXBorder = XMinMapSize + (ClienteWidth \ 2)
MaxXBorder = XMaxMapSize - (ClienteWidth \ 2)
MinYBorder = YMinMapSize + (ClienteHeight \ 2)
MaxYBorder = YMaxMapSize - (ClienteHeight \ 2)

MainViewWidth = (TilePixelWidth * WindowTileWidth)
MainViewHeight = (TilePixelHeight * WindowTileHeight)

'Resize mapdata array
ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock

'****** INIT DirectDraw ******
' Create the root DirectDraw object

Set DirectX = New DirectX7
Set DirectDraw = DirectX.DirectDrawCreate("")

DirectDraw.SetCooperativeLevel DisplayFormhWnd, DDSCL_NORMAL

'Primary Surface
With SurfaceDesc
    .lFlags = DDSD_CAPS
    .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
End With

Set PrimarySurface = DirectDraw.CreateSurface(SurfaceDesc)

Set PrimaryClipper = DirectDraw.CreateClipper(0)
PrimaryClipper.SetHWnd frmMain.hWnd
PrimarySurface.SetClipper PrimaryClipper

Set SecundaryClipper = DirectDraw.CreateClipper(0)

'Back Buffer Surface
With BackBufferRect
    .Left = 0
    .Top = 0
    .Right = TilePixelWidth * (WindowTileWidth + (2 * TileBufferSize))
    .Bottom = TilePixelHeight * (WindowTileHeight + (2 * TileBufferSize))
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

' Create surface
Set BackBufferSurface = DirectDraw.CreateSurface(SurfaceDesc)

'Set color key
ddck.low = 0
ddck.high = 0
BackBufferSurface.SetColorKey DDCKEY_SRCBLT, ddck

'Load graphic data into memory
modIndices.CargarIndicesDeGraficos

If LenB(Dir(DirGraficos & "1.bmp", vbArchive)) = 0 Then
    MsgBox "La carpeta de Graficos esta vacia o incompleta!", vbCritical
    End
Else
    frmCargando.X.Caption = "Iniciando Control de Superficies..."
'    DoEvents
    Call SurfaceDB.Initialize(DirectDraw, ClientSetup.bUseVideo, DirGraficos, ClientSetup.byMemory)
End If

'Wave Sound
Set DirectSound = DirectX.DirectSoundCreate("")
DirectSound.SetCooperativeLevel DisplayFormhWnd, DSSCL_PRIORITY
LastSoundBufferUsed = 1

InitTileEngine = True
EngineRun = True
DoEvents
End Function
