Attribute VB_Name = "modGeneral"
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
' modGeneral
'
' @remarks Funciones Generales
' @author unkwown
' @version 0.4.11
' @date 20061015

Option Explicit

Public Type typDevMODE
    dmDeviceName       As String * 32
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
    dmFormName         As String * 32
    dmUnusedPadding    As Integer
    dmBitsPerPel       As Integer
    dmPelsWidth        As Long
    dmPelsHeight       As Long
    dmDisplayFlags     As Long
    dmDisplayFrequency As Long
End Type
Public Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lptypDevMode As Any) As Boolean
Public Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lptypDevMode As Any, ByVal dwFlags As Long) As Long

Public Const CCDEVICENAME = 32
Public Const CCFORMNAME = 32
Public Const DM_BITSPERPEL = &H40000
Public Const DM_PELSWIDTH = &H80000
Public Const DM_DISPLAYFREQUENCY = &H400000
Public Const DM_PELSHEIGHT = &H100000
Public Const CDS_UPDATEREGISTRY = &H1
Public Const CDS_TEST = &H4
Public Const DISP_CHANGE_SUCCESSFUL = 0
Public Const DISP_CHANGE_RESTART = 1

Public Windows_Temp_Dir As String

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

''
' Realiza acciones de desplasamiento segun las teclas que hallamos precionado
'

Public Sub CheckKeys()
'*************************************************
'Author: ^[GS]^
'Last modified: 01/11/08
'*************************************************

If HotKeysAllow = False Then Exit Sub
        '[Loopzer]
        If GetKeyState(vbKeyControl) < 0 Then
            If Seleccionando Then
                If GetKeyState(vbKeyC) < 0 Then CopiarSeleccion
                If GetKeyState(vbKeyX) < 0 Then CortarSeleccion
                If GetKeyState(vbKeyB) < 0 Then BlockearSeleccion
                If GetKeyState(vbKeyD) < 0 Then AccionSeleccion
            Else
                If GetKeyState(vbKeyS) < 0 Then DePegar ' GS
                If GetKeyState(vbKeyV) < 0 Then PegarSeleccion
            End If
        End If
        '[/Loopzer]
        
If GetKeyState(vbKeyUp) < 0 Then
        If UserPos.y < 1 Then Exit Sub ' 10
        If LegalPos(UserPos.X, UserPos.y - 1) And WalkMode = True Then
            If dLastWalk + 50 > GetTickCount Then Exit Sub
            UserPos.y = UserPos.y - 1
            MoveCharbyPos UserCharIndex, UserPos.X, UserPos.y
            dLastWalk = GetTickCount
        ElseIf WalkMode = False Then
            UserPos.y = UserPos.y - 1
        End If
        Call ActualizaMinimap ' Radar
        frmMain.SetFocus
        Exit Sub
    End If

    If GetKeyState(vbKeyRight) < 0 Then
        If UserPos.X > 100 Then Exit Sub ' 89
        If LegalPos(UserPos.X + 1, UserPos.y) And WalkMode = True Then
            If dLastWalk + 50 > GetTickCount Then Exit Sub
            UserPos.X = UserPos.X + 1
            MoveCharbyPos UserCharIndex, UserPos.X, UserPos.y
            dLastWalk = GetTickCount
        ElseIf WalkMode = False Then
            UserPos.X = UserPos.X + 1
        End If
        Call ActualizaMinimap ' Radar
        frmMain.SetFocus
        Exit Sub
    End If

    If GetKeyState(vbKeyDown) < 0 Then
        If UserPos.y > 100 Then Exit Sub ' 92
        If LegalPos(UserPos.X, UserPos.y + 1) And WalkMode = True Then
            If dLastWalk + 50 > GetTickCount Then Exit Sub
            UserPos.y = UserPos.y + 1
            MoveCharbyPos UserCharIndex, UserPos.X, UserPos.y
            dLastWalk = GetTickCount
        ElseIf WalkMode = False Then
            UserPos.y = UserPos.y + 1
        End If
        Call ActualizaMinimap ' Radar
        frmMain.SetFocus
        Exit Sub
    End If

    If GetKeyState(vbKeyLeft) < 0 Then
        If UserPos.X < 1 Then Exit Sub ' 12
        If LegalPos(UserPos.X - 1, UserPos.y) And WalkMode = True Then
            If dLastWalk + 50 > GetTickCount Then Exit Sub
            UserPos.X = UserPos.X - 1
            MoveCharbyPos UserCharIndex, UserPos.X, UserPos.y
            dLastWalk = GetTickCount
        ElseIf WalkMode = False Then
            UserPos.X = UserPos.X - 1
        End If
        Call ActualizaMinimap ' Radar
        frmMain.SetFocus
        Exit Sub
    End If
    

End Sub

Public Function ReadField(Pos As Integer, text As String, SepASCII As Integer) As String
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
Dim i As Integer
Dim LastPos As Integer
Dim CurChar As String * 1
Dim FieldNum As Integer
Dim Seperator As String

Seperator = Chr(SepASCII)
LastPos = 0
FieldNum = 0

For i = 1 To Len(text)
    CurChar = mid(text, i, 1)
    If CurChar = Seperator Then
        FieldNum = FieldNum + 1
        If FieldNum = Pos Then
            ReadField = mid(text, LastPos + 1, (InStr(LastPos + 1, text, Seperator, vbTextCompare) - 1) - (LastPos))
            Exit Function
        End If
        LastPos = i
    End If
Next i
FieldNum = FieldNum + 1

If FieldNum = Pos Then
    ReadField = mid(text, LastPos + 1)
End If

End Function


''
' Completa y corrije un path
'
' @param Path Especifica el path con el que se trabajara
' @return   Nos devuelve el path completado

Private Function autoCompletaPath(ByVal Path As String) As String
'*************************************************
'Author: ^[GS]^
'Last modified: 22/05/06
'*************************************************
Path = Replace(Path, "/", "\")
If Left(Path, 1) = "\" Then
    ' agrego app.path & path
    Path = App.Path & Path
End If
If Right(Path, 1) <> "\" Then
    ' me aseguro que el final sea con "\"
    Path = Path & "\"
End If
autoCompletaPath = Path
End Function

''
' Carga la configuracion del WorldEditor de datos\WorldEditor.ini
'

Private Sub CargarMapIni()
'*************************************************
'Author: ^[GS]^
'Last modified: 24/11/08
'*************************************************
On Error GoTo Fallo
Dim tStr As String
Dim Leer As New clsIniReader

IniPath = App.Path & "\"

If FileExist(IniPath & "datos\WorldEditor.ini", vbArchive) = False Then
    frmMain.mnuGuardarUltimaConfig.Checked = True
    DirGraficos = IniPath & "Recursos\"
    DirIndex = App.Path & "INIT\"
    DirDats = App.Path & "DAT\"
    MaxGrhs = 32000
    UserPos.X = 50
    UserPos.y = 50
    PantallaX = 19
    PantallaY = 22
    MsgBox "Falta el archivo 'datos\WorldEditor.ini' de configuración.", vbInformation
    Exit Sub
End If

Call Leer.Initialize(IniPath & "datos\WorldEditor.ini")

' Obj de Translado
Cfg_TrOBJ = Val(Leer.GetValue("CONFIGURACION", "ObjTranslado"))
frmMain.mnuAutoCapturarTranslados.Checked = Val(Leer.GetValue("CONFIGURACION", "AutoCapturarTrans"))
frmMain.mnuAutoCapturarSuperficie.Checked = Val(Leer.GetValue("CONFIGURACION", "AutoCapturarSup"))
frmMain.mnuUtilizarDeshacer.Checked = Val(Leer.GetValue("CONFIGURACION", "UtilizarDeshacer"))

' Guardar Ultima Configuracion
frmMain.mnuGuardarUltimaConfig.Checked = Val(Leer.GetValue("CONFIGURACION", "GuardarConfig"))

' Index
MaxGrhs = Val(GetVar(IniPath & "datos\WorldEditor.ini", "INDEX", "MaxGrhs"))
If MaxGrhs < 1 Then MaxGrhs = 15000

'Reciente
frmMain.Dialog.InitDir = Leer.GetValue("PATH", "UltimoMapa")
DirGraficos = autoCompletaPath(Leer.GetValue("PATH", "DirGraficos"))
If DirGraficos = "\" Then
    DirGraficos = IniPath & "Recursos\"
End If
DirIndex = autoCompletaPath(Leer.GetValue("PATH", "DirIndex"))
If DirIndex = "\" Then
    DirIndex = IniPath & "INIT\"
End If
If FileExist(DirIndex, vbDirectory) = False Then
    MsgBox "El directorio de Index es incorrecto", vbCritical + vbOKOnly
    End
End If
DirDats = autoCompletaPath(Leer.GetValue("PATH", "DirDats"))
If DirDats = "\" Then
    DirDats = IniPath & "DAT\"
End If
If FileExist(DirDats, vbDirectory) = False Then
    MsgBox "El directorio de Dats es incorrecto", vbCritical + vbOKOnly
    End
End If

tStr = Leer.GetValue("MOSTRAR", "LastPos") ' x-y
UserPos.X = Val(ReadField(1, tStr, Asc("-")))
UserPos.y = Val(ReadField(2, tStr, Asc("-")))
If UserPos.X < XMinMapSize Or UserPos.X > XMaxMapSize Then
    UserPos.X = 50
End If
If UserPos.y < YMinMapSize Or UserPos.y > YMaxMapSize Then
    UserPos.y = 50
End If

' Menu Mostrar
frmMain.mnuVerAutomatico.Checked = Val(Leer.GetValue("MOSTRAR", "ControlAutomatico"))
frmMain.mnuVerCapa2.Checked = Val(Leer.GetValue("MOSTRAR", "Capa2"))
frmMain.mnuVerCapa3.Checked = Val(Leer.GetValue("MOSTRAR", "Capa3"))
frmMain.mnuVerCapa4.Checked = Val(Leer.GetValue("MOSTRAR", "Capa4"))
frmMain.mnuVerTranslados.Checked = Val(Leer.GetValue("MOSTRAR", "Translados"))
frmMain.mnuVerObjetos.Checked = Val(Leer.GetValue("MOSTRAR", "Objetos"))
frmMain.mnuVerNPCs.Checked = Val(Leer.GetValue("MOSTRAR", "NPCs"))
frmMain.mnuVerTriggers.Checked = Val(Leer.GetValue("MOSTRAR", "Triggers"))
frmMain.mnuVerGrilla.Checked = Val(Leer.GetValue("MOSTRAR", "Grilla")) ' Grilla
VerGrilla = frmMain.mnuVerGrilla.Checked
frmMain.mnuVerBloqueos.Checked = Val(Leer.GetValue("MOSTRAR", "Bloqueos"))
frmMain.cVerTriggers.value = frmMain.mnuVerTriggers.Checked
frmMain.cVerBloqueos.value = frmMain.mnuVerBloqueos.Checked

' Tamaño de visualizacion
PantallaX = Val(Leer.GetValue("MOSTRAR", "PantallaX"))
PantallaY = Val(Leer.GetValue("MOSTRAR", "PantallaY"))
If PantallaX > 23 Or PantallaX <= 2 Then PantallaX = 23
If PantallaY > 32 Or PantallaY <= 2 Then PantallaY = 32

' [GS] 02/10/06
' Tamaño de visualizacion en el cliente
ClienteHeight = Val(Leer.GetValue("MOSTRAR", "ClienteHeight"))
ClienteWidth = Val(Leer.GetValue("MOSTRAR", "ClienteWidth"))
If ClienteHeight <= 0 Then ClienteHeight = 13
If ClienteWidth <= 0 Then ClienteWidth = 17

Exit Sub
Fallo:
    MsgBox "ERROR " & Err.Number & " en datos\WorldEditor.ini" & vbCrLf & Err.Description, vbCritical
    Resume Next
End Sub

Public Function TomarBPP() As Integer
    Dim ModoDeVideo As typDevMODE
    Call EnumDisplaySettings(0, -1, ModoDeVideo)
    TomarBPP = CInt(ModoDeVideo.dmBitsPerPel)
End Function
Public Sub CambioDeVideo()
'*************************************************
'Author: Loopzer
'*************************************************
Exit Sub
Dim ModoDeVideo As typDevMODE
Dim R As Long
Call EnumDisplaySettings(0, -1, ModoDeVideo)
    If ModoDeVideo.dmPelsWidth < 1024 Or ModoDeVideo.dmPelsHeight < 768 Then
        Select Case MsgBox("La aplicacion necesita una resolucion minima de 1024 X 768 ,¿Acepta el Cambio de resolucion?", vbInformation + vbOKCancel, "World Editor")
            Case vbOK
                ModoDeVideo.dmPelsWidth = 1024
                ModoDeVideo.dmPelsHeight = 768
                ModoDeVideo.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
                R = ChangeDisplaySettings(ModoDeVideo, CDS_TEST)
                If R <> 0 Then
                    MsgBox "Error al cambiar la resolucion, La aplicacion se cerrara."
                    End
                End If
            Case vbCancel
                End
        End Select
    End If
End Sub
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
Public Sub Main()
'*************************************************
'Author: Unkwown
'Last modified: 25/11/08 - GS
'*************************************************
On Error Resume Next

Set Light = New clsLight

If App.PrevInstance = True Then End

Call CargarMapIni
Call IniciarCabecera(MiCabecera)

'Set Temporal Dir
Windows_Temp_Dir = General_Get_Temp_Dir

If FileExist(App.Path & "\datos\WorldEditor.jpg", vbArchive) Then frmCargando.picture1.Picture = LoadPicture(App.Path & "\datos\WorldEditor.jpg")
frmCargando.verX = "v" & App.Major & "." & App.Minor & "." & App.Revision
frmCargando.Show
frmCargando.SetFocus
DoEvents
frmCargando.X.Caption = "Iniciando DirectSound..."

DoEvents
frmCargando.X.Caption = "Cargando Indice de Superficies..."
modIndices.CargarIndicesSuperficie
DoEvents
frmCargando.X.Caption = "Indexando Cargado de Imagenes..."
LoadGrhData
CargarParticulas
CargarFxs
CargarCuerpos
DoEvents
    frmCargando.P1.Visible = True
    frmCargando.L(0).Visible = True
    frmCargando.X.Caption = "Cargando Cuerpos..."
   CargarCuerpos
    DoEvents
    frmCargando.P2.Visible = True
    frmCargando.L(1).Visible = True
    frmCargando.X.Caption = "Cargando Cabezas..."
  CargarCabezas
    DoEvents
    frmCargando.P3.Visible = True
    frmCargando.L(2).Visible = True
    frmCargando.X.Caption = "Cargando NPC's..."
    modIndices.CargarIndicesNPC
    DoEvents
    frmCargando.P4.Visible = True
    frmCargando.L(3).Visible = True
    frmCargando.X.Caption = "Cargando Objetos..."
    modIndices.CargarIndicesOBJ
    DoEvents
    frmCargando.P5.Visible = True
    frmCargando.L(4).Visible = True
    frmCargando.X.Caption = "Cargando Triggers..."
    modIndices.CargarIndicesTriggers
    DoEvents
    frmCargando.P6.Visible = True
    frmCargando.L(5).Visible = True
    DoEvents
'End If
frmCargando.SetFocus
frmCargando.X.Caption = "Iniciando Ventana de Edición..."
DoEvents
frmCargando.Hide
frmMain.Show
DoEvents
engine.Engine_Init
DoEvents
modMapIO.NuevoMapa
prgRun = True
engine.Font_Create "Tahoma", 8, False, False
engine.Start

End Sub

Public Function GetVar(file As String, Main As String, var As String) As String
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
Dim L As Integer
Dim Char As String
Dim sSpaces As String ' This will hold the input that the program will retrieve
Dim szReturn As String ' This will be the defaul value if the string is not found
szReturn = vbNullString
sSpaces = Space(5000) ' This tells the computer how long the longest string can be. If you want, you can change the number 75 to any number you wish
GetPrivateProfileString Main, var, szReturn, sSpaces, Len(sSpaces), file
GetVar = RTrim(sSpaces)
GetVar = Left(GetVar, Len(GetVar) - 1)
End Function

Public Sub WriteVar(file As String, Main As String, var As String, value As String)
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
writeprivateprofilestring Main, var, value, file
End Sub

Public Sub ToggleWalkMode()
'*************************************************
'Author: Unkwown
'Last modified: 28/05/06 - GS
'*************************************************
On Error GoTo fin:
If WalkMode = False Then
    WalkMode = True
Else
    frmMain.mnuModoCaminata.Checked = False
    WalkMode = False
End If

If WalkMode = False Then
    'Erase character
    Call EraseChar(UserCharIndex)
    MapData(UserPos.X, UserPos.y).CharIndex = 0
Else
    'MakeCharacter
    If LegalPos(UserPos.X, UserPos.y) Then
        Call MakeChar(NextOpenChar(), 1, 1, SOUTH, UserPos.X, UserPos.y)
        UserCharIndex = MapData(UserPos.X, UserPos.y).CharIndex
        frmMain.mnuModoCaminata.Checked = True
    Else
        MsgBox "ERROR: Ubicacion ilegal."
        WalkMode = False
    End If
End If
fin:
End Sub

Public Sub FixCoasts(ByVal GrhIndex As Integer, ByVal X As Integer, ByVal y As Integer)
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************

If GrhIndex = 7284 Or GrhIndex = 7290 Or GrhIndex = 7291 Or GrhIndex = 7297 Or _
   GrhIndex = 7300 Or GrhIndex = 7301 Or GrhIndex = 7302 Or GrhIndex = 7303 Or _
   GrhIndex = 7304 Or GrhIndex = 7306 Or GrhIndex = 7308 Or GrhIndex = 7310 Or _
   GrhIndex = 7311 Or GrhIndex = 7313 Or GrhIndex = 7314 Or GrhIndex = 7315 Or _
   GrhIndex = 7316 Or GrhIndex = 7317 Or GrhIndex = 7319 Or GrhIndex = 7321 Or _
   GrhIndex = 7325 Or GrhIndex = 7326 Or GrhIndex = 7327 Or GrhIndex = 7328 Or GrhIndex = 7332 Or _
   GrhIndex = 7338 Or GrhIndex = 7339 Or GrhIndex = 7345 Or GrhIndex = 7348 Or _
   GrhIndex = 7349 Or GrhIndex = 7350 Or GrhIndex = 7351 Or GrhIndex = 7352 Or _
   GrhIndex = 7349 Or GrhIndex = 7350 Or GrhIndex = 7351 Or _
   GrhIndex = 7354 Or GrhIndex = 7357 Or GrhIndex = 7358 Or GrhIndex = 7360 Or _
   GrhIndex = 7362 Or GrhIndex = 7363 Or GrhIndex = 7365 Or GrhIndex = 7366 Or _
   GrhIndex = 7367 Or GrhIndex = 7368 Or GrhIndex = 7369 Or GrhIndex = 7371 Or _
   GrhIndex = 7373 Or GrhIndex = 7375 Or GrhIndex = 7376 Then MapData(X, y).Graphic(2).GrhIndex = 0

End Sub

Public Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
Randomize Timer
RandomNumber = (UpperBound - LowerBound + 1) * Rnd + LowerBound
End Function


''
' Actualiza el Caption del menu principal
'
' @param Trabajando Indica el path del mapa con el que se esta trabajando
' @param Editado Indica si el mapa esta editado

Public Sub CaptionWorldEditor(ByVal Trabajando As String, ByVal Editado As Boolean)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If Trabajando = vbNullString Then
    Trabajando = "Nuevo Mapa"
End If
frmMain.Caption = "WorldEditor v" & App.Major & "." & App.Minor & " Build " & App.Revision & " - [" & Trabajando & "]"
If Editado = True Then
    frmMain.Caption = frmMain.Caption & " (modificado)"
End If
End Sub

Private Sub LoadClientSetup()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 26/05/2006
'26/05/2005 - GS . DirIndex
'**************************************************************
    Dim fHandle As Integer
    
    fHandle = FreeFile
    Open DirIndex & "ao.dat" For Binary Access Read Lock Write As fHandle
        Get fHandle, , ClientSetup
    Close fHandle

End Sub
Public Sub CargarParticulas()
'*************************************
'Coded by OneZero (onezero_ss@hotmail.com)
'Last Modified: 6/4/03
'Loads the Particles.ini file to the ComboBox
'Edited by Juan Martín Sotuyo Dodero to add speed and life
'*************************************
    Dim loopc As Long
    Dim i As Long
    Dim GrhListing As String
    Dim TempSet As String
    Dim ColorSet As Long
    Dim myBuffer() As Byte
    Dim StreamFile As String
    Dim Leer As New clsIniReader
    
    StreamFile = DirIndex & "Particles.ini"
    
    Leer.Initialize StreamFile
    
    TotalStreams = Val(Leer.GetValue("INIT", "Total"))
    'resize StreamData array
    ReDim StreamData(1 To TotalStreams) As Stream
    
    'fill StreamData array with info from Particles.ini
    For loopc = 1 To TotalStreams
        StreamData(loopc).name = Leer.GetValue(Val(loopc), "Name")
        frmMain.lstParticle.AddItem loopc & "-" & StreamData(loopc).name
        StreamData(loopc).NumOfParticles = Leer.GetValue(Val(loopc), "NumOfParticles")
        StreamData(loopc).X1 = Leer.GetValue(Val(loopc), "X1")
        StreamData(loopc).Y1 = Leer.GetValue(Val(loopc), "Y1")
        StreamData(loopc).X2 = Leer.GetValue(Val(loopc), "X2")
        StreamData(loopc).Y2 = Leer.GetValue(Val(loopc), "Y2")
        StreamData(loopc).angle = Leer.GetValue(Val(loopc), "Angle")
        StreamData(loopc).vecx1 = Leer.GetValue(Val(loopc), "VecX1")
        StreamData(loopc).vecx2 = Leer.GetValue(Val(loopc), "VecX2")
        StreamData(loopc).vecy1 = Leer.GetValue(Val(loopc), "VecY1")
        StreamData(loopc).vecy2 = Leer.GetValue(Val(loopc), "VecY2")
        StreamData(loopc).life1 = Leer.GetValue(Val(loopc), "Life1")
        StreamData(loopc).life2 = Leer.GetValue(Val(loopc), "Life2")
        StreamData(loopc).friction = Leer.GetValue(Val(loopc), "Friction")
        StreamData(loopc).spin = Leer.GetValue(Val(loopc), "Spin")
        StreamData(loopc).spin_speedL = Leer.GetValue(Val(loopc), "Spin_SpeedL")
        StreamData(loopc).spin_speedH = Leer.GetValue(Val(loopc), "Spin_SpeedH")
        StreamData(loopc).AlphaBlend = Leer.GetValue(Val(loopc), "AlphaBlend")
        StreamData(loopc).gravity = Leer.GetValue(Val(loopc), "Gravity")
        StreamData(loopc).grav_strength = Leer.GetValue(Val(loopc), "Grav_Strength")
        StreamData(loopc).bounce_strength = Leer.GetValue(Val(loopc), "Bounce_Strength")
        StreamData(loopc).XMove = Leer.GetValue(Val(loopc), "XMove")
        StreamData(loopc).YMove = Leer.GetValue(Val(loopc), "YMove")
        StreamData(loopc).move_x1 = Leer.GetValue(Val(loopc), "move_x1")
        StreamData(loopc).move_x2 = Leer.GetValue(Val(loopc), "move_x2")
        StreamData(loopc).move_y1 = Leer.GetValue(Val(loopc), "move_y1")
        StreamData(loopc).move_y2 = Leer.GetValue(Val(loopc), "move_y2")
        StreamData(loopc).life_counter = Leer.GetValue(Val(loopc), "life_counter")
        StreamData(loopc).speed = Val(Leer.GetValue(Val(loopc), "Speed"))
        
        StreamData(loopc).NumGrhs = Leer.GetValue(Val(loopc), "NumGrhs")
        
        ReDim StreamData(loopc).grh_list(1 To StreamData(loopc).NumGrhs)
        GrhListing = Leer.GetValue(Val(loopc), "Grh_List")
        
        For i = 1 To StreamData(loopc).NumGrhs
            StreamData(loopc).grh_list(i) = ReadField(Str(i), GrhListing, Asc(","))
        Next i
        StreamData(loopc).grh_list(i - 1) = StreamData(loopc).grh_list(i - 1)
        For ColorSet = 1 To 4
            TempSet = Leer.GetValue(Val(loopc), "ColorSet" & ColorSet)
            StreamData(loopc).colortint(ColorSet - 1).R = ReadField(1, TempSet, Asc(","))
            StreamData(loopc).colortint(ColorSet - 1).G = ReadField(2, TempSet, Asc(","))
            StreamData(loopc).colortint(ColorSet - 1).B = ReadField(3, TempSet, Asc(","))
        Next ColorSet
    Next loopc
    
End Sub
 
Public Function General_Var_Get(ByVal file As String, ByVal Main As String, ByVal var As String) As String
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Get a var to from a text file
'*****************************************************************
    Dim L As Long
    Dim Char As String
    Dim sSpaces As String 'Input that the program will retrieve
    Dim szReturn As String 'Default value if the string is not found
   
    szReturn = ""
   
    sSpaces = Space$(5000)
   
    GetPrivateProfileString Main, var, szReturn, sSpaces, Len(sSpaces), file
   
    General_Var_Get = RTrim$(sSpaces)
    General_Var_Get = Left$(General_Var_Get, Len(General_Var_Get) - 1)
End Function

Public Function General_Field_Read(ByVal field_pos As Long, ByVal text As String, ByVal delimiter As Byte) As String
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
    For i = 1 To Len(text)
        If delimiter = CByte(Asc(mid$(text, i, 1))) Then
            FieldNum = FieldNum + 1
            If FieldNum = field_pos Then
                General_Field_Read = mid$(text, LastPos + 1, (InStr(LastPos + 1, text, Chr$(delimiter), vbTextCompare) - 1) - (LastPos))
                Exit Function
            End If
            LastPos = i
        End If
    Next i
    FieldNum = FieldNum + 1
    If FieldNum = field_pos Then
        General_Field_Read = mid$(text, LastPos + 1)
    End If
End Function

Public Sub FillListFiles(ListBox As ListBox, Extension As String, Path As String)
    On Error GoTo ErrFLF
   
    Dim tFile As file
    Dim tFolder As Folder
   
    'Limpiamos el ListBox.
    ListBox.Clear
   
    Set tFolder = FSO.GetFolder(Path)
   
    'Listamos los ficheros de la carpeta seteada en tFolder.
    For Each tFile In tFolder.Files
        'Filtramos y listamos.
        If Right$(tFile.name, 4) = "." & Extension Then
            ListBox.AddItem tFile.name
        End If
    Next tFile
   
    Exit Sub
ErrFLF:
    MsgBox "Error. " & Err.Description & "(" & Err.Number & ")", vbCritical, "Error"
End Sub
'**************************************************************
Public Sub DibujarMiniMapa()
Dim map_x, map_y, Capas As Byte
Dim loopc As Long
    For map_y = 1 To 100
        For map_x = 1 To 100
        For Capas = 1 To 2
            If MapData(map_x, map_y).Graphic(Capas).GrhIndex > 0 Then
                SetPixel frmMain.Minimap.hdc, map_x - 1, map_y - 1, GrhData(MapData(map_x, map_y).Graphic(Capas).GrhIndex).MiniMap_color
            End If
            If MapData(map_x, map_y).Graphic(4).GrhIndex > 0 And VerCapa4 And Not bTecho Then
                SetPixel frmMain.Minimap.hdc, map_x - 1, map_y - 1, GrhData(MapData(map_x, map_y).Graphic(4).GrhIndex).MiniMap_color
            End If
        Next Capas
        Next map_x
    Next map_y
    
    For loopc = 1 To LastChar
    If charlist(loopc).active = 1 Then
        MapData(charlist(loopc).Pos.X, charlist(loopc).Pos.y).CharIndex = loopc
        If charlist(loopc).Heading <> 0 Then
            SetPixel frmMain.Minimap.hdc, 0 + charlist(loopc).Pos.X, 0 + charlist(loopc).Pos.y, RGB(0, 255, 0)
            SetPixel frmMain.Minimap.hdc, 0 + charlist(loopc).Pos.X, 1 + charlist(loopc).Pos.y, RGB(0, 255, 0)
        End If
    End If
Next loopc
   
    frmMain.Minimap.Refresh
End Sub
Public Sub ActualizaMinimap()
    frmMain.UserArea.Left = UserPos.X - 9
    frmMain.UserArea.Top = UserPos.y - 8
End Sub
'***********************************************************
Private Function LoadGrhData() As Boolean
On Error GoTo ErrorHandler
    Dim Grh As Long
    Dim Frame As Long
    Dim grhCount As Long
    Dim handle As Integer
    Dim fileVersion As Long
   
    'Open files
    handle = FreeFile()

    Open App.Path & "\INIT\Graficos.ind" For Binary Access Read As handle
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
            
           GrhData(Grh).active = True

            ReDim .Frames(1 To GrhData(Grh).NumFrames)
           
            If .NumFrames > 1 Then
                'Read a animation GRH set
                For Frame = 1 To .NumFrames
                    Get handle, , .Frames(Frame)
                    If .Frames(Frame) <= 0 Or .Frames(Frame) > grhCount Then
                        GoTo ErrorHandler
                    End If
                Next Frame
               
                Get handle, , .speed
               
                If .speed <= 0 Then GoTo ErrorHandler
               
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
   
Dim Count As Long
 
Open App.Path & "\INIT\minimap.dat" For Binary As #1
    Seek #1, 1
    For Count = 1 To grhCount
        If GrhData(Count).active Then
            Get #1, , GrhData(Count).MiniMap_color
        End If
    Next Count
Close #1
 
    LoadGrhData = True
Exit Function
 
ErrorHandler:
    LoadGrhData = False
End Function
