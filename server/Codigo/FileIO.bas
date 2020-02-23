Attribute VB_Name = "ES"
'Winter-AO 0.12.2
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Winter-AO is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Public Sub CargarSpawnList()
    Dim N As Integer, LoopC As Integer
    N = val(GetVar(App.Path & "\Dat\Invokar.dat", "INIT", "NumNPCs"))
    ReDim SpawnList(N) As tCriaturasEntrenador
    For LoopC = 1 To N
        SpawnList(LoopC).NpcIndex = val(GetVar(App.Path & "\Dat\Invokar.dat", "LIST", "NI" & LoopC))
        SpawnList(LoopC).NpcName = GetVar(App.Path & "\Dat\Invokar.dat", "LIST", "NN" & LoopC)
    Next LoopC
    
End Sub

Function EsAdmin(ByVal Name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim NomB As String

NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "Admines"))

For WizNum = 1 To NumWizs
    NomB = UCase$(GetVar(IniPath & "Server.ini", "Admines", "Admin" & WizNum))
    
    If Left$(NomB, 1) = "*" Or Left$(NomB, 1) = "+" Then NomB = Right$(NomB, Len(NomB) - 1)
    If UCase$(Name) = NomB Then
        EsAdmin = True
        Exit Function
    End If
Next WizNum
EsAdmin = False

End Function

Function EsDios(ByVal Name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim NomB As String

NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "Dioses"))
For WizNum = 1 To NumWizs
    NomB = UCase$(GetVar(IniPath & "Server.ini", "Dioses", "Dios" & WizNum))
    
    If Left$(NomB, 1) = "*" Or Left$(NomB, 1) = "+" Then NomB = Right$(NomB, Len(NomB) - 1)
    If UCase$(Name) = NomB Then
        EsDios = True
        Exit Function
    End If
Next WizNum
EsDios = False
End Function

Function EsSemiDios(ByVal Name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim NomB As String

NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "SemiDioses"))
For WizNum = 1 To NumWizs
    NomB = UCase$(GetVar(IniPath & "Server.ini", "SemiDioses", "SemiDios" & WizNum))
    
    If Left$(NomB, 1) = "*" Or Left$(NomB, 1) = "+" Then NomB = Right$(NomB, Len(NomB) - 1)
    If UCase$(Name) = NomB Then
        EsSemiDios = True
        Exit Function
    End If
Next WizNum
EsSemiDios = False

End Function

Function EsConsejero(ByVal Name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim NomB As String

NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "Consejeros"))
For WizNum = 1 To NumWizs
    NomB = UCase$(GetVar(IniPath & "Server.ini", "Consejeros", "Consejero" & WizNum))
    
    If Left$(NomB, 1) = "*" Or Left$(NomB, 1) = "+" Then NomB = Right$(NomB, Len(NomB) - 1)
    If UCase$(Name) = NomB Then
        EsConsejero = True
        Exit Function
    End If
Next WizNum
EsConsejero = False
End Function

Function EsRolesMaster(ByVal Name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim NomB As String

NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "RolesMasters"))
For WizNum = 1 To NumWizs
    NomB = UCase$(GetVar(IniPath & "Server.ini", "RolesMasters", "RM" & WizNum))
    
    If Left$(NomB, 1) = "*" Or Left$(NomB, 1) = "+" Then NomB = Right$(NomB, Len(NomB) - 1)
    If UCase$(Name) = NomB Then
        EsRolesMaster = True
        Exit Function
    End If
Next WizNum
EsRolesMaster = False
End Function


Public Function TxtDimension(ByVal Name As String) As Long
Dim N As Integer, cad As String, Tam As Long
N = FreeFile(1)
Open Name For Input As #N
Tam = 0
Do While Not EOF(N)
    Tam = Tam + 1
    Line Input #N, cad
Loop
Close N
TxtDimension = Tam
End Function

Public Sub CargarForbidenWords()

ReDim ForbidenNames(1 To TxtDimension(DatPath & "NombresInvalidos.txt"))
Dim N As Integer, i As Integer
N = FreeFile(1)
Open DatPath & "NombresInvalidos.txt" For Input As #N

For i = 1 To UBound(ForbidenNames)
    Line Input #N, ForbidenNames(i)
Next i

Close N

End Sub

Public Sub CargarHechizos()

'###################################################
'#               ATENCION PELIGRO                  #
'###################################################
'
'  ¡¡¡¡ NO USAR GetVar PARA LEER Hechizos.dat !!!!
'
'El que ose desafiar esta LEY, se las tendrá que ver
'con migo. Para leer Hechizos.dat se deberá usar
'la nueva clase clsLeerInis.
'
'Alejo
'
'###################################################

On Error GoTo Errhandler

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando Hechizos."

Dim Hechizo As Integer
Dim Leer As New clsIniReader

Call Leer.Initialize(DatPath & "Hechizos.dat")

'obtiene el numero de hechizos
NumeroHechizos = val(Leer.GetValue("INIT", "NumeroHechizos"))

ReDim Hechizos(1 To NumeroHechizos) As tHechizo

frmCargando.cargar.min = 0
frmCargando.cargar.max = NumeroHechizos
frmCargando.cargar.Value = 0

'Llena la lista
For Hechizo = 1 To NumeroHechizos
    With Hechizos(Hechizo)
        .Nombre = Leer.GetValue("Hechizo" & Hechizo, "Nombre")
        .desc = Leer.GetValue("Hechizo" & Hechizo, "Desc")
        .PalabrasMagicas = Leer.GetValue("Hechizo" & Hechizo, "PalabrasMagicas")
        
        .HechizeroMsg = Leer.GetValue("Hechizo" & Hechizo, "HechizeroMsg")
        .TargetMsg = Leer.GetValue("Hechizo" & Hechizo, "TargetMsg")
        .PropioMsg = Leer.GetValue("Hechizo" & Hechizo, "PropioMsg")
        
        .Tipo = val(Leer.GetValue("Hechizo" & Hechizo, "Tipo"))
        .WAV = val(Leer.GetValue("Hechizo" & Hechizo, "WAV"))
        .FXgrh = val(Leer.GetValue("Hechizo" & Hechizo, "Fxgrh"))
        .Particle = val(Leer.GetValue("Hechizo" & Hechizo, "Particle"))
        
        .loops = val(Leer.GetValue("Hechizo" & Hechizo, "Loops"))
        
    '    .Resis = val(Leer.GetValue("Hechizo" & Hechizo, "Resis"))
        
        .SubeHP = val(Leer.GetValue("Hechizo" & Hechizo, "SubeHP"))
        .MinHP = val(Leer.GetValue("Hechizo" & Hechizo, "MinHP"))
        .MaxHP = val(Leer.GetValue("Hechizo" & Hechizo, "MaxHP"))
        
        .SubeMana = val(Leer.GetValue("Hechizo" & Hechizo, "SubeMana"))
        .MiMana = val(Leer.GetValue("Hechizo" & Hechizo, "MinMana"))
        .MaMana = val(Leer.GetValue("Hechizo" & Hechizo, "MaxMana"))
        
        .SubeSta = val(Leer.GetValue("Hechizo" & Hechizo, "SubeSta"))
        .MinSta = val(Leer.GetValue("Hechizo" & Hechizo, "MinSta"))
        .MaxSta = val(Leer.GetValue("Hechizo" & Hechizo, "MaxSta"))
        
        .SubeHam = val(Leer.GetValue("Hechizo" & Hechizo, "SubeHam"))
        .MinHam = val(Leer.GetValue("Hechizo" & Hechizo, "MinHam"))
        .MaxHam = val(Leer.GetValue("Hechizo" & Hechizo, "MaxHam"))
        
        .SubeSed = val(Leer.GetValue("Hechizo" & Hechizo, "SubeSed"))
        .MinSed = val(Leer.GetValue("Hechizo" & Hechizo, "MinSed"))
        .MaxSed = val(Leer.GetValue("Hechizo" & Hechizo, "MaxSed"))
        
        .SubeAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "SubeAG"))
        .MinAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "MinAG"))
        .MaxAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "MaxAG"))
        
        .SubeFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "SubeFU"))
        .MinFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "MinFU"))
        .MaxFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "MaxFU"))
        
        .SubeCarisma = val(Leer.GetValue("Hechizo" & Hechizo, "SubeCA"))
        .MinCarisma = val(Leer.GetValue("Hechizo" & Hechizo, "MinCA"))
        .MaxCarisma = val(Leer.GetValue("Hechizo" & Hechizo, "MaxCA"))
        
        .Area = val(Leer.GetValue("Hechizo" & Hechizo, "Area"))
        
        .Invisibilidad = val(Leer.GetValue("Hechizo" & Hechizo, "Invisibilidad"))
        .Paraliza = val(Leer.GetValue("Hechizo" & Hechizo, "Paraliza"))
        .Inmoviliza = val(Leer.GetValue("Hechizo" & Hechizo, "Inmoviliza"))
        .RemoverParalisis = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverParalisis"))
        .RemoverEstupidez = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverEstupidez"))
        .RemueveInvisibilidadParcial = val(Leer.GetValue("Hechizo" & Hechizo, "RemueveInvisibilidadParcial"))
        
        
        .CuraVeneno = val(Leer.GetValue("Hechizo" & Hechizo, "CuraVeneno"))
        .Envenena = val(Leer.GetValue("Hechizo" & Hechizo, "Envenena"))
        .Maldicion = val(Leer.GetValue("Hechizo" & Hechizo, "Maldicion"))
        .RemoverMaldicion = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverMaldicion"))
        .Bendicion = val(Leer.GetValue("Hechizo" & Hechizo, "Bendicion"))
        .Revivir = val(Leer.GetValue("Hechizo" & Hechizo, "Revivir"))
        
        .Ceguera = val(Leer.GetValue("Hechizo" & Hechizo, "Ceguera"))
        .Estupidez = val(Leer.GetValue("Hechizo" & Hechizo, "Estupidez"))
        
        .Invoca = val(Leer.GetValue("Hechizo" & Hechizo, "Invoca"))
        .NumNpc = val(Leer.GetValue("Hechizo" & Hechizo, "NumNpc"))
        .cant = val(Leer.GetValue("Hechizo" & Hechizo, "Cant"))
        .Mimetiza = val(Leer.GetValue("hechizo" & Hechizo, "Mimetiza"))
        
        
    '    .Materializa = val(Leer.GetValue("Hechizo" & Hechizo, "Materializa"))
    '    .ItemIndex = val(Leer.GetValue("Hechizo" & Hechizo, "ItemIndex"))
        
        .MinSkill = val(Leer.GetValue("Hechizo" & Hechizo, "MinSkill"))
        .ManaRequerido = val(Leer.GetValue("Hechizo" & Hechizo, "ManaRequerido"))
        
        'Barrin 30/9/03
        .StaRequerido = val(Leer.GetValue("Hechizo" & Hechizo, "StaRequerido"))
        
        .Target = val(Leer.GetValue("Hechizo" & Hechizo, "Target"))
        frmCargando.cargar.Value = frmCargando.cargar.Value + 1
        
        .NeedStaff = val(Leer.GetValue("Hechizo" & Hechizo, "NeedStaff"))
        .StaffAffected = CBool(val(Leer.GetValue("Hechizo" & Hechizo, "StaffAffected")))
        
        .ExclusivoClase = UCase$(Leer.GetValue("Hechizo" & Hechizo, "ExclusivoClase"))
    End With
Next Hechizo

Set Leer = Nothing
Exit Sub

Errhandler:
 MsgBox "Error cargando hechizos.dat " & Err.Number & ": " & Err.description
 
End Sub

Sub LoadMotd()
Dim i As Integer

MaxLines = val(GetVar(App.Path & "\Dat\Motd.ini", "INIT", "NumLines"))

ReDim MOTD(1 To MaxLines)
For i = 1 To MaxLines
    MOTD(i).texto = GetVar(App.Path & "\Dat\Motd.ini", "Motd", "Line" & i)
    MOTD(i).Formato = vbNullString
Next i

End Sub

Public Sub DarPremioCastillos()
On Error GoTo handler
Dim LoopC As Integer
    For LoopC = 1 To LastUser
        If UserList(LoopC).GuildIndex <> 0 Then
            If modGuilds.GuildName(UserList(LoopC).GuildIndex) = CastilloWINTER Then
                UserList(LoopC).Stats.GLD = UserList(LoopC).Stats.GLD + 10000
                Call WriteConsoleMsg((LoopC), "Has recibido 10000 monedas de oro por mantener el castillo Winter", FontTypeNames.FONTTYPE_GLD)
                Call WriteUpdateUserStats(LoopC)
                'Call EnviarDatosASlot(LoopC, PrepareMessagePlayWave(43, .Pos.X, .Pos.Y))
            End If
        End If
    Next LoopC
Exit Sub
handler:
Call LogError("Error en DarPremioCastillos.")
End Sub

Public Sub DoBackUp()
'Call LogTarea("Sub DoBackUp")
haciendoBK = True
Dim i As Integer
Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())

Call LimpiarMundo
Call WorldSave
Call modGuilds.v_RutinaElecciones
Call ResetCentinelaInfo     'Reseteamos al centinela

Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())

haciendoBK = False

'Log
On Error Resume Next
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\BackUps.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time
Close #nfile
End Sub

Public Sub GrabarMapa(ByVal map As Long, ByRef MAPFILE As String)
'***************************************************
'Author: Unknown
'Last Modification: 12/01/2011
'10/08/2010 - Pato: Implemento el clsByteBuffer para el grabado de mapas
'28/10/2010:ZaMa - Ahora no se hace backup de los pretorianos.
'12/01/2011 - Amraphen: Ahora no se hace backup de NPCs prohibidos (Pretorianos, Mascotas, Invocados y Centinela)
'***************************************************

On Error Resume Next
    Dim FreeFileMap As Long
    Dim FreeFileInf As Long
    Dim Y As Long
    Dim X As Long
    Dim ByFlags As Byte
    Dim LoopC As Long
    Dim MapWriter As clsByteBuffer
    Dim InfWriter As clsByteBuffer
    Dim IniManager As clsIniManager
    Dim NpcInvalido As Boolean
    
    Set MapWriter = New clsByteBuffer
    Set InfWriter = New clsByteBuffer
    Set IniManager = New clsIniManager
    
    If FileExist(MAPFILE & ".map", vbNormal) Then
        Kill MAPFILE & ".map"
    End If
    
    If FileExist(MAPFILE & ".inf", vbNormal) Then
        Kill MAPFILE & ".inf"
    End If
    
    'Open .map file
    FreeFileMap = FreeFile
    Open MAPFILE & ".Map" For Binary As FreeFileMap
    
    Call MapWriter.initializeWriter(FreeFileMap)
    
    'Open .inf file
    FreeFileInf = FreeFile
    Open MAPFILE & ".Inf" For Binary As FreeFileInf
    
    Call InfWriter.initializeWriter(FreeFileInf)
    
    'map Header
    Call MapWriter.putInteger(MapInfo(map).MapVersion)
        
    Call MapWriter.putString(MiCabecera.desc, False)
    Call MapWriter.putLong(MiCabecera.crc)
    Call MapWriter.putLong(MiCabecera.MagicWord)
    
    Call MapWriter.putDouble(0)
    
    'inf Header
    Call InfWriter.putDouble(0)
    Call InfWriter.putInteger(0)
    
    'Write .map file
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            With MapData(map, X, Y)
                ByFlags = 0
                
                If .Blocked Then ByFlags = ByFlags Or 1
                If .Graphic(2) Then ByFlags = ByFlags Or 2
                If .Graphic(3) Then ByFlags = ByFlags Or 4
                If .Graphic(4) Then ByFlags = ByFlags Or 8
                If .trigger Then ByFlags = ByFlags Or 16
                
                Call MapWriter.putByte(ByFlags)
                
                Call MapWriter.putInteger(.Graphic(1))
                
                For LoopC = 2 To 4
                    If .Graphic(LoopC) Then _
                        Call MapWriter.putInteger(.Graphic(LoopC))
                Next LoopC
                
                If .trigger Then _
                    Call MapWriter.putInteger(CInt(.trigger))
                
                '.inf file
                ByFlags = 0
                
                If .ObjInfo.ObjIndex > 0 Then
                   If ObjData(.ObjInfo.ObjIndex).OBJType = eOBJType.otFogata Then
                        .ObjInfo.ObjIndex = 0
                        .ObjInfo.amount = 0
                    End If
                End If
    
                If .TileExit.map Then ByFlags = ByFlags Or 1

                If Not NpcInvalido Then ByFlags = ByFlags Or 2
                
                If .ObjInfo.ObjIndex Then ByFlags = ByFlags Or 4
                
                Call InfWriter.putByte(ByFlags)
                
                If .TileExit.map Then
                    Call InfWriter.putInteger(.TileExit.map)
                    Call InfWriter.putInteger(.TileExit.X)
                    Call InfWriter.putInteger(.TileExit.Y)
                End If
                
                If .NpcIndex And Not NpcInvalido Then _
                    Call InfWriter.putInteger(Npclist(.NpcIndex).Numero)
                
                If .ObjInfo.ObjIndex Then
                    Call InfWriter.putInteger(.ObjInfo.ObjIndex)
                    Call InfWriter.putInteger(.ObjInfo.amount)
                End If
                
                NpcInvalido = False
            End With
        Next X
    Next Y
    
    Call MapWriter.saveBuffer
    Call InfWriter.saveBuffer
    
    'Close .map file
    Close FreeFileMap

    'Close .inf file
    Close FreeFileInf
    
    Set MapWriter = Nothing
    Set InfWriter = Nothing

    With MapInfo(map)
        'write .dat file
        Call IniManager.ChangeValue("Mapa" & map, "Name", .Name)
        Call IniManager.ChangeValue("Mapa" & map, "MusicNum", .Music)
        Call IniManager.ChangeValue("Mapa" & map, "MagiaSinefecto", .MagiaSinEfecto)
        Call IniManager.ChangeValue("Mapa" & map, "InviSinEfecto", .InviSinEfecto)
        Call IniManager.ChangeValue("Mapa" & map, "ResuSinEfecto", .ResuSinEfecto)
        Call IniManager.ChangeValue("Mapa" & map, "StartPos", .StartPos.map & "-" & .StartPos.X & "-" & .StartPos.Y)

    
        Call IniManager.ChangeValue("Mapa" & map, "Terreno", TerrainByteToString(.Terreno))
        Call IniManager.ChangeValue("Mapa" & map, "Zona", .Zona)
        Call IniManager.ChangeValue("Mapa" & map, "Restringir", RestrictByteToString(.Restringir))
        Call IniManager.ChangeValue("Mapa" & map, "BackUp", str(.BackUp))
    
        If .Pk Then
            Call IniManager.ChangeValue("Mapa" & map, "Pk", "0")
        Else
            Call IniManager.ChangeValue("Mapa" & map, "Pk", "1")
        End If
        
        Call IniManager.ChangeValue("Mapa" & map, "NoEncriptarMP", .NoEncriptarMP)
        Call IniManager.ChangeValue("Mapa" & map, "lvlminimo", .lvlMinimo)
        
        Call IniManager.DumpFile(MAPFILE & ".dat")
    End With
    
    Set IniManager = Nothing
End Sub
Sub LoadArmasHerreria()

Dim N As Integer, lc As Integer

N = val(GetVar(DatPath & "ArmasHerrero.dat", "INIT", "NumArmas"))

ReDim Preserve ArmasHerrero(1 To N) As Integer

For lc = 1 To N
    ArmasHerrero(lc) = val(GetVar(DatPath & "ArmasHerrero.dat", "Arma" & lc, "Index"))
Next lc

End Sub

Sub LoadArmadurasHerreria()

Dim N As Integer, lc As Integer

N = val(GetVar(DatPath & "ArmadurasHerrero.dat", "INIT", "NumArmaduras"))

ReDim Preserve ArmadurasHerrero(1 To N) As Integer

For lc = 1 To N
    ArmadurasHerrero(lc) = val(GetVar(DatPath & "ArmadurasHerrero.dat", "Armadura" & lc, "Index"))
Next lc

End Sub

Sub LoadBalance()
    Dim i As Long
    
    'Modificadores de Clase
    For i = 1 To NUMCLASES
        ModClase(i).Evasion = val(GetVar(DatPath & "Balance.dat", "MODEVASION", ListaClases(i)))
        ModClase(i).AtaqueArmas = val(GetVar(DatPath & "Balance.dat", "MODATAQUEARMAS", ListaClases(i)))
        ModClase(i).AtaqueProyectiles = val(GetVar(DatPath & "Balance.dat", "MODATAQUEPROYECTILES", ListaClases(i)))
        ModClase(i).DañoArmas = val(GetVar(DatPath & "Balance.dat", "MODDAÑOARMAS", ListaClases(i)))
        ModClase(i).DañoProyectiles = val(GetVar(DatPath & "Balance.dat", "MODDAÑOPROYECTILES", ListaClases(i)))
        ModClase(i).DañoWrestling = val(GetVar(DatPath & "Balance.dat", "MODDAÑOWRESTLING", ListaClases(i)))
        ModClase(i).Escudo = val(GetVar(DatPath & "Balance.dat", "MODESCUDO", ListaClases(i)))
    Next i
    
    'Modificadores de Raza
    For i = 1 To NUMRAZAS
        ModRaza(i).Fuerza = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Fuerza"))
        ModRaza(i).Agilidad = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Agilidad"))
        ModRaza(i).Inteligencia = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Inteligencia"))
        ModRaza(i).Carisma = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Carisma"))
        ModRaza(i).Constitucion = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Constitucion"))
    Next i
    
    'Modificadores de Vida
    For i = 1 To NUMCLASES
        ModVida(i) = val(GetVar(DatPath & "Balance.dat", "MODVIDA", ListaClases(i)))
    Next i
    
    'Distribución de Vida
    For i = 1 To 5
        DistribucionEnteraVida(i) = val(GetVar(DatPath & "Balance.dat", "DISTRIBUCION", "E" + CStr(i)))
    Next i
    For i = 1 To 4
        DistribucionSemienteraVida(i) = val(GetVar(DatPath & "Balance.dat", "DISTRIBUCION", "S" + CStr(i)))
    Next i
    
    'Extra
    PorcentajeRecuperoMana = val(GetVar(DatPath & "Balance.dat", "EXTRA", "PorcentajeRecuperoMana"))

    'Party
    ExponenteNivelParty = val(GetVar(DatPath & "Balance.dat", "PARTY", "ExponenteNivelParty"))
End Sub

Sub LoadObjCarpintero()

Dim N As Integer, lc As Integer

N = val(GetVar(DatPath & "ObjCarpintero.dat", "INIT", "NumObjs"))

ReDim Preserve ObjCarpintero(1 To N) As Integer

For lc = 1 To N
    ObjCarpintero(lc) = val(GetVar(DatPath & "ObjCarpintero.dat", "Obj" & lc, "Index"))
Next lc

End Sub

Public Sub LoadCanjes()

Dim N As Integer, lc As Integer

N = val(GetVar(DatPath & "ObjCanjes.dat", "INIT", "NumObjs"))

ReDim Preserve ObjCanjes(1 To N) As Integer

For lc = 1 To N
    ObjCanjes(lc) = val(GetVar(DatPath & "ObjCanjes.dat", "Obj" & lc, "Index"))
Next lc

End Sub

Sub LoadOBJData()

'###################################################
'#               ATENCION PELIGRO                  #
'###################################################
'
'¡¡¡¡ NO USAR GetVar PARA LEER DESDE EL OBJ.DAT !!!!
'
'El que ose desafiar esta LEY, se las tendrá que ver
'con migo. Para leer desde el OBJ.DAT se deberá usar
'la nueva clase clsLeerInis.
'
'Alejo
'
'###################################################

'Call LogTarea("Sub LoadOBJData")

On Error GoTo Errhandler

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando base de datos de los objetos."

'*****************************************************************
'Carga la lista de objetos
'*****************************************************************
Dim Object As Integer
Dim Leer As New clsIniReader

Call Leer.Initialize(DatPath & "Obj.dat")

'obtiene el numero de obj
NumObjDatas = val(Leer.GetValue("INIT", "NumObjs"))

frmCargando.cargar.min = 0
frmCargando.cargar.max = NumObjDatas
frmCargando.cargar.Value = 0


ReDim Preserve ObjData(1 To NumObjDatas) As ObjData
  
'Llena la lista
For Object = 1 To NumObjDatas
        
    ObjData(Object).Name = Leer.GetValue("OBJ" & Object, "Name")
    
    'Pablo (ToxicWaste) Log de Objetos.
    ObjData(Object).Log = val(Leer.GetValue("OBJ" & Object, "Log"))
    ObjData(Object).NoLog = val(Leer.GetValue("OBJ" & Object, "NoLog"))
    '07/09/07
    
    ObjData(Object).GrhIndex = val(Leer.GetValue("OBJ" & Object, "GrhIndex"))
    If ObjData(Object).GrhIndex = 0 Then
        ObjData(Object).GrhIndex = ObjData(Object).GrhIndex
    End If
    
    ObjData(Object).OBJType = val(Leer.GetValue("OBJ" & Object, "ObjType"))
    
    ObjData(Object).Newbie = val(Leer.GetValue("OBJ" & Object, "Newbie"))
    ObjData(Object).Aura = val(Leer.GetValue("OBJ" & Object, "CreaAura"))
    
    Select Case ObjData(Object).OBJType
        Case eOBJType.otArmadura
            ObjData(Object).Real = val(Leer.GetValue("OBJ" & Object, "Real"))
            ObjData(Object).Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
            ObjData(Object).LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
        
        Case eOBJType.otESCUDO
            ObjData(Object).ShieldAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
            ObjData(Object).LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
            ObjData(Object).Real = val(Leer.GetValue("OBJ" & Object, "Real"))
            ObjData(Object).Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
        
        Case eOBJType.otCASCO
            ObjData(Object).CascoAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
            ObjData(Object).LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
            ObjData(Object).Real = val(Leer.GetValue("OBJ" & Object, "Real"))
            ObjData(Object).Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
        
        Case eOBJType.otWeapon
            ObjData(Object).WeaponAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
            ObjData(Object).Apuñala = val(Leer.GetValue("OBJ" & Object, "Apuñala"))
            ObjData(Object).Envenena = val(Leer.GetValue("OBJ" & Object, "Envenena"))
            ObjData(Object).MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
            ObjData(Object).PROYECTIL = val(Leer.GetValue("OBJ" & Object, "Proyectil"))
            ObjData(Object).Municion = val(Leer.GetValue("OBJ" & Object, "Municiones"))
            ObjData(Object).StaffPower = val(Leer.GetValue("OBJ" & Object, "StaffPower"))
            ObjData(Object).StaffDamageBonus = val(Leer.GetValue("OBJ" & Object, "StaffDamageBonus"))
            ObjData(Object).Refuerzo = val(Leer.GetValue("OBJ" & Object, "Refuerzo"))
            
            ObjData(Object).LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
            ObjData(Object).Real = val(Leer.GetValue("OBJ" & Object, "Real"))
            ObjData(Object).Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
            ObjData(Object).VaraDragon = val(Leer.GetValue("OBJ" & Object, "VaraDragon"))
 
        Case eOBJType.otInstrumentos
            ObjData(Object).Snd1 = val(Leer.GetValue("OBJ" & Object, "SND1"))
            ObjData(Object).Snd2 = val(Leer.GetValue("OBJ" & Object, "SND2"))
            ObjData(Object).Snd3 = val(Leer.GetValue("OBJ" & Object, "SND3"))
            'Pablo (ToxicWaste)
            ObjData(Object).Real = val(Leer.GetValue("OBJ" & Object, "Real"))
            ObjData(Object).Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
        
        Case eOBJType.otMinerales
            ObjData(Object).MinSkill = val(Leer.GetValue("OBJ" & Object, "MinSkill"))
        
        Case eOBJType.otPuertas, eOBJType.otBotellaVacia, eOBJType.otBotellaLlena
            ObjData(Object).IndexAbierta = val(Leer.GetValue("OBJ" & Object, "IndexAbierta"))
            ObjData(Object).IndexCerrada = val(Leer.GetValue("OBJ" & Object, "IndexCerrada"))
            ObjData(Object).IndexCerradaLlave = val(Leer.GetValue("OBJ" & Object, "IndexCerradaLlave"))
        
        Case otPociones
            ObjData(Object).TipoPocion = val(Leer.GetValue("OBJ" & Object, "TipoPocion"))
            ObjData(Object).MaxModificador = val(Leer.GetValue("OBJ" & Object, "MaxModificador"))
            ObjData(Object).MinModificador = val(Leer.GetValue("OBJ" & Object, "MinModificador"))
            ObjData(Object).DuracionEfecto = val(Leer.GetValue("OBJ" & Object, "DuracionEfecto"))
        
        Case eOBJType.otBarcos
            ObjData(Object).MinSkill = val(Leer.GetValue("OBJ" & Object, "MinSkill"))
            ObjData(Object).MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
        
        Case eOBJType.otMonturas
            ObjData(Object).MinSkill = val(Leer.GetValue("OBJ" & Object, "MinSkill"))
            ObjData(Object).MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
        
        Case eOBJType.otFlechas
            ObjData(Object).MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
            ObjData(Object).Envenena = val(Leer.GetValue("OBJ" & Object, "Envenena"))
            ObjData(Object).Paraliza = val(Leer.GetValue("OBJ" & Object, "Paraliza"))
            
        Case eOBJType.otAnillo 'Pablo (ToxicWaste)
            ObjData(Object).LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
            
        Case eOBJType.otPasajes
            ObjData(Object).DesdeMap = val(Leer.GetValue("OBJ" & Object, "DesdeMap"))
            ObjData(Object).HastaMap = val(Leer.GetValue("OBJ" & Object, "HastaMap"))
            ObjData(Object).HastaX = val(Leer.GetValue("OBJ" & Object, "HastaX"))
            ObjData(Object).HastaY = val(Leer.GetValue("OBJ" & Object, "HastaY"))
            
        Case eOBJType.otMetaPot
            ObjData(Object).MetaModificador = val(Leer.GetValue("OBJ" & Object, "MetaModificador"))
            
    End Select
    
    ObjData(Object).Ropaje = val(Leer.GetValue("OBJ" & Object, "NumRopaje"))
    ObjData(Object).HechizoIndex = val(Leer.GetValue("OBJ" & Object, "HechizoIndex"))
    
    ObjData(Object).Meta = val(Leer.GetValue("OBJ" & Object, "Meta")) 'Metamorfosis
    
    ObjData(Object).LingoteIndex = val(Leer.GetValue("OBJ" & Object, "LingoteIndex"))
    
    ObjData(Object).MineralIndex = val(Leer.GetValue("OBJ" & Object, "MineralIndex"))
    
    ObjData(Object).MaxHP = val(Leer.GetValue("OBJ" & Object, "MaxHP"))
    ObjData(Object).MinHP = val(Leer.GetValue("OBJ" & Object, "MinHP"))
    
    ObjData(Object).Mujer = val(Leer.GetValue("OBJ" & Object, "Mujer"))
    ObjData(Object).Hombre = val(Leer.GetValue("OBJ" & Object, "Hombre"))
    
    ObjData(Object).MinHam = val(Leer.GetValue("OBJ" & Object, "MinHam"))
    ObjData(Object).MinSed = val(Leer.GetValue("OBJ" & Object, "MinAgu"))
    
    ObjData(Object).MinDef = val(Leer.GetValue("OBJ" & Object, "MINDEF"))
    ObjData(Object).MaxDef = val(Leer.GetValue("OBJ" & Object, "MAXDEF"))
    ObjData(Object).def = (ObjData(Object).MinDef + ObjData(Object).MaxDef) / 2
    
    ObjData(Object).RazaEnana = val(Leer.GetValue("OBJ" & Object, "RazaEnana"))
    ObjData(Object).RazaDrow = val(Leer.GetValue("OBJ" & Object, "RazaDrow"))
    ObjData(Object).RazaElfa = val(Leer.GetValue("OBJ" & Object, "RazaElfa"))
    ObjData(Object).RazaGnoma = val(Leer.GetValue("OBJ" & Object, "RazaGnoma"))
    ObjData(Object).RazaHumana = val(Leer.GetValue("OBJ" & Object, "RazaHumana"))
    ObjData(Object).RazaOrca = val(Leer.GetValue("OBJ" & Object, "RazaOrca"))
    
    ObjData(Object).Valor = val(Leer.GetValue("OBJ" & Object, "Valor"))
    
    ObjData(Object).Crucial = val(Leer.GetValue("OBJ" & Object, "Crucial"))
    
    ObjData(Object).Cerrada = val(Leer.GetValue("OBJ" & Object, "abierta"))
    If ObjData(Object).Cerrada = 1 Then
        ObjData(Object).Llave = val(Leer.GetValue("OBJ" & Object, "Llave"))
        ObjData(Object).clave = val(Leer.GetValue("OBJ" & Object, "Clave"))
    End If
    
    'Puertas y llaves
    ObjData(Object).clave = val(Leer.GetValue("OBJ" & Object, "Clave"))
    
    ObjData(Object).texto = Leer.GetValue("OBJ" & Object, "Texto")
    ObjData(Object).GrhSecundario = val(Leer.GetValue("OBJ" & Object, "VGrande"))
    
    ObjData(Object).Agarrable = val(Leer.GetValue("OBJ" & Object, "Agarrable"))
    ObjData(Object).ForoID = Leer.GetValue("OBJ" & Object, "ID")
    
    
    'CHECK: !!! Esto es provisorio hasta que los de Dateo cambien los valores de string a numerico
    Dim i As Integer
    Dim N As Integer
    Dim S As String
    For i = 1 To NUMCLASES
        S = UCase$(Leer.GetValue("OBJ" & Object, "CP" & i))
        N = 1
        Do While LenB(S) > 0 And UCase$(ListaClases(N)) <> S
            N = N + 1
        Loop
        ObjData(Object).ClaseProhibida(i) = IIf(LenB(S) > 0, N, 0)
    Next i
    
    ObjData(Object).DefensaMagicaMax = val(Leer.GetValue("OBJ" & Object, "DefensaMagicaMax"))
    ObjData(Object).DefensaMagicaMin = val(Leer.GetValue("OBJ" & Object, "DefensaMagicaMin"))
    
    ObjData(Object).SkCarpinteria = val(Leer.GetValue("OBJ" & Object, "SkCarpinteria"))
        
    If ObjData(Object).SkCarpinteria > 0 Then _
        ObjData(Object).Madera = val(Leer.GetValue("OBJ" & Object, "Madera"))
    
    'Bebidas
    ObjData(Object).MinSta = val(Leer.GetValue("OBJ" & Object, "MinST"))
    
    ObjData(Object).NoSeCae = val(Leer.GetValue("OBJ" & Object, "NoSeCae"))
    
    ObjData(Object).Canjeable = val(Leer.GetValue("OBJ" & Object, "Canjeable"))
    ObjData(Object).Puntos = val(Leer.GetValue("OBJ" & Object, "Puntos"))
    
    frmCargando.cargar.Value = frmCargando.cargar.Value + 1
Next Object

Set Leer = Nothing

Exit Sub

Errhandler:
    MsgBox "error cargando objetos " & Err.Number & ": " & Err.description


End Sub

Sub LoadUserStats(ByVal UserIndex As Integer, ByRef UserFile As clsIniReader)

Dim LoopC As Long

For LoopC = 1 To NUMATRIBUTOS
  UserList(UserIndex).Stats.UserAtributos(LoopC) = CInt(UserFile.GetValue("ATRIBUTOS", "AT" & LoopC))
  UserList(UserIndex).Stats.UserAtributosBackUP(LoopC) = UserList(UserIndex).Stats.UserAtributos(LoopC)
Next LoopC

For LoopC = 1 To NUMSKILLS
  UserList(UserIndex).Stats.UserSkills(LoopC) = CInt(UserFile.GetValue("SKILLS", "SK" & LoopC))
Next LoopC

For LoopC = 1 To MAXUSERHECHIZOS
  UserList(UserIndex).Stats.UserHechizos(LoopC) = CInt(UserFile.GetValue("Hechizos", "H" & LoopC))
Next LoopC


UserList(UserIndex).Stats.GLD = CLng(UserFile.GetValue("STATS", "GLD"))
UserList(UserIndex).Stats.Banco = CLng(UserFile.GetValue("STATS", "BANCO"))

UserList(UserIndex).Stats.MaxHP = CInt(UserFile.GetValue("STATS", "MaxHP"))
UserList(UserIndex).Stats.MinHP = CInt(UserFile.GetValue("STATS", "MinHP"))

UserList(UserIndex).Stats.MinSta = CInt(UserFile.GetValue("STATS", "MinSTA"))
UserList(UserIndex).Stats.MaxSta = CInt(UserFile.GetValue("STATS", "MaxSTA"))

UserList(UserIndex).Stats.MaxMAN = CInt(UserFile.GetValue("STATS", "MaxMAN"))
UserList(UserIndex).Stats.MinMAN = CInt(UserFile.GetValue("STATS", "MinMAN"))

UserList(UserIndex).Stats.MaxHIT = CInt(UserFile.GetValue("STATS", "MaxHIT"))
UserList(UserIndex).Stats.MinHIT = CInt(UserFile.GetValue("STATS", "MinHIT"))

UserList(UserIndex).Stats.MaxAGU = CByte(UserFile.GetValue("STATS", "MaxAGU"))
UserList(UserIndex).Stats.MinAGU = CByte(UserFile.GetValue("STATS", "MinAGU"))

UserList(UserIndex).Stats.MaxHam = CByte(UserFile.GetValue("STATS", "MaxHAM"))
UserList(UserIndex).Stats.MinHam = CByte(UserFile.GetValue("STATS", "MinHAM"))

UserList(UserIndex).Stats.SkillPts = CInt(UserFile.GetValue("STATS", "SkillPtsLibres"))

UserList(UserIndex).Stats.Exp = CDbl(UserFile.GetValue("STATS", "EXP"))
UserList(UserIndex).Stats.ELU = CLng(UserFile.GetValue("STATS", "ELU"))
UserList(UserIndex).Stats.ELV = CByte(UserFile.GetValue("STATS", "ELV"))
UserList(UserIndex).Stats.Puntos = CByte(UserFile.GetValue("STATS", "Puntos"))

UserList(UserIndex).Stats.DañoExtra = 0

UserList(UserIndex).Stats.UsuariosMatados = CLng(UserFile.GetValue("MUERTES", "UserMuertes"))
UserList(UserIndex).Stats.NPCsMuertos = CInt(UserFile.GetValue("MUERTES", "NpcsMuertes"))

If CByte(UserFile.GetValue("CONSEJO", "PERTENECE")) Then _
    UserList(UserIndex).flags.Privilegios = UserList(UserIndex).flags.Privilegios Or PlayerType.RoyalCouncil

If CByte(UserFile.GetValue("CONSEJO", "PERTENECECAOS")) Then _
    UserList(UserIndex).flags.Privilegios = UserList(UserIndex).flags.Privilegios Or PlayerType.ChaosCouncil

End Sub

Sub LoadUserReputacion(ByVal UserIndex As Integer, ByRef UserFile As clsIniReader)

UserList(UserIndex).Reputacion.AsesinoRep = val(UserFile.GetValue("REP", "Asesino"))
UserList(UserIndex).Reputacion.BandidoRep = val(UserFile.GetValue("REP", "Bandido"))
UserList(UserIndex).Reputacion.BurguesRep = val(UserFile.GetValue("REP", "Burguesia"))
UserList(UserIndex).Reputacion.LadronesRep = val(UserFile.GetValue("REP", "Ladrones"))
UserList(UserIndex).Reputacion.NobleRep = val(UserFile.GetValue("REP", "Nobles"))
UserList(UserIndex).Reputacion.PlebeRep = val(UserFile.GetValue("REP", "Plebe"))
UserList(UserIndex).Reputacion.Promedio = val(UserFile.GetValue("REP", "Promedio"))

End Sub

Sub LoadUserInit(ByVal UserIndex As Integer, ByRef UserFile As clsIniReader)
'*************************************************
'Author: Unknown
'Last modified: 19/11/2006
'Loads the Users records
'23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
'23/01/2007 Pablo (ToxicWaste) - Quito CriminalesMatados de Stats porque era redundante.
'*************************************************
Dim LoopC As Long
Dim ln As String

UserList(UserIndex).Faccion.ArmadaReal = CByte(UserFile.GetValue("FACCIONES", "EjercitoReal"))
UserList(UserIndex).Faccion.FuerzasCaos = CByte(UserFile.GetValue("FACCIONES", "EjercitoCaos"))
UserList(UserIndex).Faccion.CiudadanosMatados = CLng(UserFile.GetValue("FACCIONES", "CiudMatados"))
UserList(UserIndex).Faccion.CriminalesMatados = CLng(UserFile.GetValue("FACCIONES", "CrimMatados"))
UserList(UserIndex).Faccion.RecibioArmaduraCaos = CByte(UserFile.GetValue("FACCIONES", "rArCaos"))
UserList(UserIndex).Faccion.RecibioArmaduraReal = CByte(UserFile.GetValue("FACCIONES", "rArReal"))
UserList(UserIndex).Faccion.RecibioExpInicialCaos = CByte(UserFile.GetValue("FACCIONES", "rExCaos"))
UserList(UserIndex).Faccion.RecibioExpInicialReal = CByte(UserFile.GetValue("FACCIONES", "rExReal"))
UserList(UserIndex).Faccion.RecompensasCaos = CLng(UserFile.GetValue("FACCIONES", "recCaos"))
UserList(UserIndex).Faccion.RecompensasReal = CLng(UserFile.GetValue("FACCIONES", "recReal"))
UserList(UserIndex).Faccion.Reenlistadas = CByte(UserFile.GetValue("FACCIONES", "Reenlistadas"))
UserList(UserIndex).Faccion.NivelIngreso = CInt(UserFile.GetValue("FACCIONES", "NivelIngreso"))
UserList(UserIndex).Faccion.FechaIngreso = UserFile.GetValue("FACCIONES", "FechaIngreso")
UserList(UserIndex).Faccion.MatadosIngreso = CInt(UserFile.GetValue("FACCIONES", "MatadosIngreso"))
UserList(UserIndex).Faccion.NextRecompensa = CInt(UserFile.GetValue("FACCIONES", "NextRecompensa"))

UserList(UserIndex).flags.Muerto = CByte(UserFile.GetValue("FLAGS", "Muerto"))
UserList(UserIndex).flags.Escondido = CByte(UserFile.GetValue("FLAGS", "Escondido"))

UserList(UserIndex).flags.Hambre = CByte(UserFile.GetValue("FLAGS", "Hambre"))
UserList(UserIndex).flags.Sed = CByte(UserFile.GetValue("FLAGS", "Sed"))
UserList(UserIndex).flags.Desnudo = CByte(UserFile.GetValue("FLAGS", "Desnudo"))

UserList(UserIndex).flags.Navegando = CByte(UserFile.GetValue("FLAGS", "Navegando"))
UserList(UserIndex).flags.Equitando = CByte(UserFile.GetValue("FLAGS", "Equitando"))

UserList(UserIndex).flags.Metamorfosis = CByte(UserFile.GetValue("FLAGS", "Metamorfosis")) 'Metamorfosis

UserList(UserIndex).flags.Envenenado = CByte(UserFile.GetValue("FLAGS", "Envenenado"))
UserList(UserIndex).flags.Paralizado = CByte(UserFile.GetValue("FLAGS", "Paralizado"))
If UserList(UserIndex).flags.Paralizado = 1 Then
    UserList(UserIndex).Counters.Paralisis = IntervaloParalizado
End If


UserList(UserIndex).Counters.Pena = CLng(UserFile.GetValue("COUNTERS", "Pena"))
UserList(UserIndex).Counters.Piedra = CLng(UserFile.GetValue("COUNTERS", "Piedra"))

UserList(UserIndex).email = UserFile.GetValue("CONTACTO", "Email")

UserList(UserIndex).genero = UserFile.GetValue("INIT", "Genero")
UserList(UserIndex).clase = UserFile.GetValue("INIT", "Clase")
UserList(UserIndex).raza = UserFile.GetValue("INIT", "Raza")
UserList(UserIndex).Hogar = UserFile.GetValue("INIT", "Hogar")
UserList(UserIndex).Char.heading = CInt(UserFile.GetValue("INIT", "Heading"))


UserList(UserIndex).OrigChar.Head = CInt(UserFile.GetValue("INIT", "Head"))
UserList(UserIndex).OrigChar.body = CInt(UserFile.GetValue("INIT", "Body"))
UserList(UserIndex).OrigChar.WeaponAnim = CInt(UserFile.GetValue("INIT", "Arma"))
UserList(UserIndex).OrigChar.ShieldAnim = CInt(UserFile.GetValue("INIT", "Escudo"))
UserList(UserIndex).OrigChar.CascoAnim = CInt(UserFile.GetValue("INIT", "Casco"))
UserList(UserIndex).OrigChar.Aura = CInt(UserFile.GetValue("INIT", "Aura"))

#If ConUpTime Then
    UserList(UserIndex).UpTime = CLng(UserFile.GetValue("INIT", "UpTime"))
#End If

UserList(UserIndex).OrigChar.heading = eHeading.SOUTH

If UserList(UserIndex).flags.Muerto = 0 Then
    UserList(UserIndex).Char = UserList(UserIndex).OrigChar
Else
    UserList(UserIndex).Char.body = iCuerpoMuerto
    UserList(UserIndex).Char.Head = iCabezaMuerto
    UserList(UserIndex).Char.WeaponAnim = NingunArma
    UserList(UserIndex).Char.ShieldAnim = NingunEscudo
    UserList(UserIndex).Char.CascoAnim = NingunCasco
    UserList(UserIndex).Char.CascoAnim = 0
End If


UserList(UserIndex).desc = UserFile.GetValue("INIT", "Desc")

UserList(UserIndex).Pos.map = CInt(ReadField(1, UserFile.GetValue("INIT", "Position"), 45))
UserList(UserIndex).Pos.X = CInt(ReadField(2, UserFile.GetValue("INIT", "Position"), 45))
UserList(UserIndex).Pos.Y = CInt(ReadField(3, UserFile.GetValue("INIT", "Position"), 45))

UserList(UserIndex).Invent.NroItems = CInt(UserFile.GetValue("Inventory", "CantidadItems"))

'[KEVIN]--------------------------------------------------------------------
'***********************************************************************************
UserList(UserIndex).BancoInvent.NroItems = CInt(UserFile.GetValue("BancoInventory", "CantidadItems"))
'Lista de objetos del banco
For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
    ln = UserFile.GetValue("BancoInventory", "Obj" & LoopC)
    UserList(UserIndex).BancoInvent.Object(LoopC).ObjIndex = CInt(ReadField(1, ln, 45))
    UserList(UserIndex).BancoInvent.Object(LoopC).amount = CInt(ReadField(2, ln, 45))
Next LoopC
'------------------------------------------------------------------------------------
'[/KEVIN]*****************************************************************************


'Lista de objetos
For LoopC = 1 To MAX_INVENTORY_SLOTS
    ln = UserFile.GetValue("Inventory", "Obj" & LoopC)
    UserList(UserIndex).Invent.Object(LoopC).ObjIndex = CInt(ReadField(1, ln, 45))
    UserList(UserIndex).Invent.Object(LoopC).amount = CInt(ReadField(2, ln, 45))
    UserList(UserIndex).Invent.Object(LoopC).Equipped = CByte(ReadField(3, ln, 45))
Next LoopC

'Obtiene el indice-objeto del arma
UserList(UserIndex).Invent.WeaponEqpSlot = CByte(UserFile.GetValue("Inventory", "WeaponEqpSlot"))
If UserList(UserIndex).Invent.WeaponEqpSlot > 0 Then
    UserList(UserIndex).Invent.WeaponEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.WeaponEqpSlot).ObjIndex
End If

'Obtiene el indice-objeto del armadura
UserList(UserIndex).Invent.ArmourEqpSlot = CByte(UserFile.GetValue("Inventory", "ArmourEqpSlot"))
If UserList(UserIndex).Invent.ArmourEqpSlot > 0 Then
    UserList(UserIndex).Invent.ArmourEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.ArmourEqpSlot).ObjIndex
    UserList(UserIndex).flags.Desnudo = 0
Else
    UserList(UserIndex).flags.Desnudo = 1
End If

'Obtiene el indice-objeto del escudo
UserList(UserIndex).Invent.EscudoEqpSlot = CByte(UserFile.GetValue("Inventory", "EscudoEqpSlot"))
If UserList(UserIndex).Invent.EscudoEqpSlot > 0 Then
    UserList(UserIndex).Invent.EscudoEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.EscudoEqpSlot).ObjIndex
End If

'Obtiene el indice-objeto del casco
UserList(UserIndex).Invent.CascoEqpSlot = CByte(UserFile.GetValue("Inventory", "CascoEqpSlot"))
If UserList(UserIndex).Invent.CascoEqpSlot > 0 Then
    UserList(UserIndex).Invent.CascoEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.CascoEqpSlot).ObjIndex
End If

'Obtiene el indice-objeto barco
UserList(UserIndex).Invent.BarcoSlot = CByte(UserFile.GetValue("Inventory", "BarcoSlot"))
If UserList(UserIndex).Invent.BarcoSlot > 0 Then
    UserList(UserIndex).Invent.BarcoObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.BarcoSlot).ObjIndex
End If

'Obtiene el indice-objeto monturas
UserList(UserIndex).Invent.MonturaSlot = CByte(UserFile.GetValue("Inventory", "MonturaSlot"))
If UserList(UserIndex).Invent.MonturaSlot > 0 Then
    UserList(UserIndex).Invent.MonturaObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.MonturaSlot).ObjIndex
End If

'Obtiene el indice-objeto municion
UserList(UserIndex).Invent.MunicionEqpSlot = CByte(UserFile.GetValue("Inventory", "MunicionSlot"))
If UserList(UserIndex).Invent.MunicionEqpSlot > 0 Then
    UserList(UserIndex).Invent.MunicionEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.MunicionEqpSlot).ObjIndex
End If

'[Alejo]
'Obtiene el indice-objeto anilo
UserList(UserIndex).Invent.AnilloEqpSlot = CByte(UserFile.GetValue("Inventory", "AnilloSlot"))
If UserList(UserIndex).Invent.AnilloEqpSlot > 0 Then
    UserList(UserIndex).Invent.AnilloEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.AnilloEqpSlot).ObjIndex
End If

UserList(UserIndex).NroMascotas = CInt(UserFile.GetValue("MASCOTAS", "NroMascotas"))
Dim NpcIndex As Integer
For LoopC = 1 To MAXMASCOTAS
    UserList(UserIndex).MascotasType(LoopC) = val(UserFile.GetValue("MASCOTAS", "MAS" & LoopC))
Next LoopC

ln = UserFile.GetValue("Guild", "GUILDINDEX")
If IsNumeric(ln) Then
    UserList(UserIndex).GuildIndex = CInt(ln)
Else
    UserList(UserIndex).GuildIndex = 0
End If

End Sub

Function GetVar(ByVal File As String, ByVal Main As String, ByVal Var As String, Optional EmptySpaces As Long = 1024) As String

Dim sSpaces As String ' This will hold the input that the program will retrieve
Dim szReturn As String ' This will be the defaul value if the string is not found
  
szReturn = vbNullString
  
sSpaces = Space$(EmptySpaces) ' This tells the computer how long the longest string can be
  
  
GetPrivateProfileString Main, Var, szReturn, sSpaces, EmptySpaces, File
  
GetVar = RTrim$(sSpaces)
GetVar = Left$(GetVar, Len(GetVar) - 1)
  
End Function

Sub CargarBackUp()

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando backup."

Dim map As Integer
Dim TempInt As Integer
Dim tFileName As String
Dim npcfile As String

On Error GoTo man
    
    NumMaps = val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
    Call InitAreas
    
    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumMaps
    frmCargando.cargar.Value = 0
    
    MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")
    
    
    ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    ReDim MapInfo(1 To NumMaps) As MapInfo
    
    For map = 1 To NumMaps
        If val(GetVar(App.Path & MapPath & "Mapa" & map & ".Dat", "Mapa" & map, "BackUp")) <> 0 Then
            tFileName = App.Path & "\WorldBackUp\Mapa" & map
        Else
            tFileName = App.Path & MapPath & "Mapa" & map
        End If
        
        Call CargarMapa(map, tFileName)
        
        frmCargando.cargar.Value = frmCargando.cargar.Value + 1
        DoEvents
    Next map

Exit Sub

man:
    MsgBox ("Error durante la carga de mapas, el mapa " & map & " contiene errores")
    Call LogError(Date & " " & Err.description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.source)
 
End Sub

Sub LoadMapData()

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando mapas..."

Dim map As Integer
Dim TempInt As Integer
Dim tFileName As String
Dim npcfile As String

On Error GoTo man
    
    NumMaps = val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
    Call InitAreas
    
    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumMaps
    frmCargando.cargar.Value = 0
    
    MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")
    
    
    ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    ReDim MapInfo(1 To NumMaps) As MapInfo
      
    For map = 1 To NumMaps
        
        tFileName = App.Path & MapPath & "Mapa" & map
        Call CargarMapa(map, tFileName)
        
        frmCargando.cargar.Value = frmCargando.cargar.Value + 1
        DoEvents
    Next map

Exit Sub

man:
    MsgBox ("Error durante la carga de mapas, el mapa " & map & " contiene errores")
    Call LogError(Date & " " & Err.description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.source)

End Sub

Public Sub CargarMapa(ByVal map As Long, ByRef MAPFl As String)
'***************************************************
'Author: Unknown
'Last Modification: 10/08/2010
'10/08/2010 - Pato: Implemento el clsByteBuffer y el clsIniManager para la carga de mapa
'***************************************************

On Error GoTo errh
    Dim hFile As Integer
    Dim X As Long
    Dim Y As Long
    Dim ByFlags As Byte
    Dim npcfile As String
    Dim Leer As clsIniManager
    Dim MapReader As clsByteBuffer
    Dim InfReader As clsByteBuffer
    Dim Buff() As Integer
    
    Set MapReader = New clsByteBuffer
    Set InfReader = New clsByteBuffer
    Set Leer = New clsIniManager
    
    npcfile = DatPath & "NPCs.dat"
    
    hFile = FreeFile

    Open MAPFl & ".map" For Binary As #hFile
        Seek hFile, 1

        ReDim Buff(LOF(hFile) - 1) As Integer
    
        Get #hFile, , Buff
    Close hFile
    
    Call MapReader.initializeReader(Buff)

    'inf
    Open MAPFl & ".inf" For Binary As #hFile
        Seek hFile, 1

        ReDim Buff(LOF(hFile) - 1) As Integer
    
        Get #hFile, , Buff
    Close hFile
    
    Call InfReader.initializeReader(Buff)
    
    'map Header
    MapInfo(map).MapVersion = MapReader.getInteger
    
    MiCabecera.desc = MapReader.getString(Len(MiCabecera.desc))
    MiCabecera.crc = MapReader.getLong
    MiCabecera.MagicWord = MapReader.getLong
    
    Call MapReader.getDouble

    'inf Header
    Call InfReader.getDouble
    Call InfReader.getInteger

    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            With MapData(map, X, Y)
                '.map file
                ByFlags = MapReader.getByte

                If ByFlags And 1 Then .Blocked = 1

                .Graphic(1) = MapReader.getInteger

                'Layer 2 used?
                If ByFlags And 2 Then .Graphic(2) = MapReader.getInteger

                'Layer 3 used?
                If ByFlags And 4 Then .Graphic(3) = MapReader.getInteger

                'Layer 4 used?
                If ByFlags And 8 Then .Graphic(4) = MapReader.getInteger

                'Trigger used?
                If ByFlags And 16 Then .trigger = MapReader.getInteger
                
                If ByFlags And 32 Then .particle_group_index = MapReader.getInteger
            
                If ByFlags And 64 Then
                    .base_light(0) = MapReader.getInteger
                    .base_light(1) = MapReader.getInteger
                    .base_light(2) = MapReader.getInteger
                    .base_light(3) = MapReader.getInteger
                    
                    If .base_light(0) Then _
                        .light_value(0) = MapReader.getLong
                    
                    If .base_light(1) Then _
                        .light_value(1) = MapReader.getLong
                    
                    If .base_light(2) Then _
                        .light_value(2) = MapReader.getLong
                    
                    If .base_light(3) Then _
                        .light_value(3) = MapReader.getLong
                End If

                '.inf file
                ByFlags = InfReader.getByte

                If ByFlags And 1 Then
                    .TileExit.map = InfReader.getInteger
                    .TileExit.X = InfReader.getInteger
                    .TileExit.Y = InfReader.getInteger
                End If

                If ByFlags And 2 Then
                    'Get and make NPC
                     .NpcIndex = InfReader.getInteger

                    If .NpcIndex > 0 Then
                        'Si el npc debe hacer respawn en la pos
                        'original la guardamos
                        If val(GetVar(npcfile, "NPC" & .NpcIndex, "PosOrig")) = 1 Then
                            .NpcIndex = OpenNPC(.NpcIndex)
                            Npclist(.NpcIndex).Orig.map = map
                            Npclist(.NpcIndex).Orig.X = X
                            Npclist(.NpcIndex).Orig.Y = Y
                        Else
                            .NpcIndex = OpenNPC(.NpcIndex)
                        End If

                        Npclist(.NpcIndex).Pos.map = map
                        Npclist(.NpcIndex).Pos.X = X
                        Npclist(.NpcIndex).Pos.Y = Y

                        Call MakeNPCChar(True, 0, .NpcIndex, map, X, Y)
                    End If
                End If

                If ByFlags And 4 Then
                    'Get and make Object
                    .ObjInfo.ObjIndex = InfReader.getInteger
                    .ObjInfo.amount = InfReader.getInteger
                End If
            End With
        Next X
    Next Y
    
    Call Leer.Initialize(MAPFl & ".dat")
    
    With MapInfo(map)
        .Name = Leer.GetValue("Mapa" & map, "Name")
        .Music = Leer.GetValue("Mapa" & map, "MusicNum")
        .StartPos.map = val(ReadField(1, Leer.GetValue("Mapa" & map, "StartPos"), Asc("-")))
        .StartPos.X = val(ReadField(2, Leer.GetValue("Mapa" & map, "StartPos"), Asc("-")))
        .StartPos.Y = val(ReadField(3, Leer.GetValue("Mapa" & map, "StartPos"), Asc("-")))
        
        .MagiaSinEfecto = val(Leer.GetValue("Mapa" & map, "MagiaSinEfecto"))
        .InviSinEfecto = val(Leer.GetValue("Mapa" & map, "InviSinEfecto"))
        .ResuSinEfecto = val(Leer.GetValue("Mapa" & map, "ResuSinEfecto"))
        
        .NoEncriptarMP = val(Leer.GetValue("Mapa" & map, "NoEncriptarMP"))
        
        If val(Leer.GetValue("Mapa" & map, "Pk")) = 0 Then
            .Pk = True
        Else
            .Pk = False
        End If
        
        .Terreno = TerrainStringToByte(Leer.GetValue("Mapa" & map, "Terreno"))
        .Zona = Leer.GetValue("Mapa" & map, "Zona")
        .Restringir = RestrictStringToByte(Leer.GetValue("Mapa" & map, "Restringir"))
        .BackUp = val(Leer.GetValue("Mapa" & map, "BACKUP"))
        .lvlMinimo = val(Leer.GetValue("Mapa" & map, "lvlminimo"))
    End With
    
    Set MapReader = Nothing
    Set InfReader = Nothing
    Set Leer = Nothing
    
    Erase Buff
Exit Sub

errh:
    Call LogError("Error cargando mapa: " & map & " - Pos: " & X & "," & Y & "." & Err.description)

    Set MapReader = Nothing
    Set InfReader = Nothing
    Set Leer = Nothing
End Sub

Sub LoadSini()

Dim Temporal As Long

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando info de inicio del server."

BootDelBackUp = val(GetVar(IniPath & "Server.ini", "INIT", "IniciarDesdeBackUp"))

CastilloWINTER = GetVar(App.Path & "\Castillos.ini", "CLANES", "WINTER")
NPCReyCastle = val(GetVar(App.Path & "\Castillos.ini", "REY", "NUMERO"))
WINTER = val(GetVar(App.Path & "\Castillos.ini", "CASTILLO", "WINTER"))

Puerto = val(GetVar(IniPath & "Server.ini", "INIT", "StartPort"))
HideMe = val(GetVar(IniPath & "Server.ini", "INIT", "Hide"))
OroX = val(GetVar(IniPath & "Server.ini", "INIT", "OroX"))
ExpX = val(GetVar(IniPath & "Server.ini", "INIT", "ExpX"))

centinelaActivado = val(GetVar(IniPath & "Server.ini", "INIT", "centinelaActivado"))

AllowMultiLogins = val(GetVar(IniPath & "Server.ini", "INIT", "AllowMultiLogins"))
IdleLimit = val(GetVar(IniPath & "Server.ini", "INIT", "IdleLimit"))
'Lee la version correcta del cliente
ULTIMAVERSION = GetVar(IniPath & "Server.ini", "INIT", "Version")

PuedeCrearPersonajes = val(GetVar(IniPath & "Server.ini", "INIT", "PuedeCrearPersonajes"))
ServerSoloGMs = val(GetVar(IniPath & "Server.ini", "init", "ServerSoloGMs"))

ArmaduraImperial1 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraImperial1"))
ArmaduraImperial2 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraImperial2"))
ArmaduraImperial3 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraImperial3"))
TunicaMagoImperial = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaMagoImperial"))
TunicaMagoImperialEnanos = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaMagoImperialEnanos"))
ArmaduraCaos1 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraCaos1"))
ArmaduraCaos2 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraCaos2"))
ArmaduraCaos3 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraCaos3"))
TunicaMagoCaos = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaMagoCaos"))
TunicaMagoCaosEnanos = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaMagoCaosEnanos"))

VestimentaImperialHumano = val(GetVar(IniPath & "Server.ini", "INIT", "VestimentaImperialHumano"))
VestimentaImperialEnano = val(GetVar(IniPath & "Server.ini", "INIT", "VestimentaImperialEnano"))
TunicaConspicuaHumano = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaConspicuaHumano"))
TunicaConspicuaEnano = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaConspicuaEnano"))
ArmaduraNobilisimaHumano = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraNobilisimaHumano"))
ArmaduraNobilisimaEnano = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraNobilisimaEnano"))
ArmaduraGranSacerdote = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraGranSacerdote"))

VestimentaLegionHumano = val(GetVar(IniPath & "Server.ini", "INIT", "VestimentaLegionHumano"))
VestimentaLegionEnano = val(GetVar(IniPath & "Server.ini", "INIT", "VestimentaLegionEnano"))
TunicaLobregaHumano = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaLobregaHumano"))
TunicaLobregaEnano = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaLobregaEnano"))
TunicaEgregiaHumano = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaEgregiaHumano"))
TunicaEgregiaEnano = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaEgregiaEnano"))
SacerdoteDemoniaco = val(GetVar(IniPath & "Server.ini", "INIT", "SacerdoteDemoniaco"))

EnTesting = val(GetVar(IniPath & "Server.ini", "INIT", "Testing"))

'Intervalos
SanaIntervaloSinDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "SanaIntervaloSinDescansar"))
FrmInterv.txtSanaIntervaloSinDescansar.Text = SanaIntervaloSinDescansar

StaminaIntervaloSinDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "StaminaIntervaloSinDescansar"))
FrmInterv.txtStaminaIntervaloSinDescansar.Text = StaminaIntervaloSinDescansar

SanaIntervaloDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "SanaIntervaloDescansar"))
FrmInterv.txtSanaIntervaloDescansar.Text = SanaIntervaloDescansar

StaminaIntervaloDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "StaminaIntervaloDescansar"))
FrmInterv.txtStaminaIntervaloDescansar.Text = StaminaIntervaloDescansar

IntervaloSed = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloSed"))
FrmInterv.txtIntervaloSed.Text = IntervaloSed

IntervaloHambre = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloHambre"))
FrmInterv.txtIntervaloHambre.Text = IntervaloHambre

IntervaloVeneno = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloVeneno"))
FrmInterv.txtIntervaloVeneno.Text = IntervaloVeneno

IntervaloParalizado = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParalizado"))
FrmInterv.txtIntervaloParalizado.Text = IntervaloParalizado

IntervaloInvisible = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvisible"))
FrmInterv.txtIntervaloInvisible.Text = IntervaloInvisible

IntervaloFrio = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloFrio"))
FrmInterv.txtIntervaloFrio.Text = IntervaloFrio

IntervaloWavFx = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloWAVFX"))
FrmInterv.txtIntervaloWAVFX.Text = IntervaloWavFx

IntervaloInvocacion = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvocacion"))
FrmInterv.txtInvocacion.Text = IntervaloInvocacion

IntervaloParaConexion = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParaConexion"))
FrmInterv.txtIntervaloParaConexion.Text = IntervaloParaConexion

IntervaloPuedeMakrear = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloMakreo"))

IntervaloPuedeResucitar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloResu"))

'&&&&&&&&&&&&&&&&&&&&& TIMERS &&&&&&&&&&&&&&&&&&&&&&&


IntervaloUserPuedeCastear = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloLanzaHechizo"))
FrmInterv.txtIntervaloLanzaHechizo.Text = IntervaloUserPuedeCastear

frmMain.TIMER_AI.Interval = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloNpcAI"))
FrmInterv.txtAI.Text = frmMain.TIMER_AI.Interval

frmMain.npcataca.Interval = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloNpcPuedeAtacar"))
FrmInterv.txtNPCPuedeAtacar.Text = frmMain.npcataca.Interval

IntervaloUserPuedeTrabajar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloTrabajo"))
FrmInterv.txtTrabajo.Text = IntervaloUserPuedeTrabajar

IntervaloUserPuedeAtacar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeAtacar"))
FrmInterv.txtPuedeAtacar.Text = IntervaloUserPuedeAtacar

'TODO : Agregar estos intervalos al form!!!
IntervaloMagiaGolpe = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloMagiaGolpe"))
IntervaloGolpeMagia = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloGolpeMagia"))
IntervaloGolpeUsar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloGolpeUsar"))

MinutosWs = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloWS"))
If MinutosWs < 60 Then MinutosWs = 180

IntervaloCerrarConexion = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloCerrarConexion"))
IntervaloUserPuedeUsar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeUsar"))
IntervaloFlechasCazadores = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloFlechasCazadores"))

IntervaloOculto = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloOculto"))

'&&&&&&&&&&&&&&&&&&&&& FIN TIMERS &&&&&&&&&&&&&&&&&&&&&&&
  
recordusuarios = val(GetVar(IniPath & "Server.ini", "INIT", "Record"))
  
'Max users
Temporal = val(GetVar(IniPath & "Server.ini", "INIT", "MaxUsers"))
If MaxUsers = 0 Then
    MaxUsers = Temporal
    ReDim UserList(1 To MaxUsers) As User
End If

'&&&&&&&&&&&&&&&&&&&&& BALANCE &&&&&&&&&&&&&&&&&&&&&&&
'Se agregó en LoadBalance y en el Balance.dat
'PorcentajeRecuperoMana = val(GetVar(IniPath & "Server.ini", "BALANCE", "PorcentajeRecuperoMana"))

''&&&&&&&&&&&&&&&&&&&&& FIN BALANCE &&&&&&&&&&&&&&&&&&&&&&&
Call Statistics.Initialize

Ramx.map = 1
Ramx.X = 63
Ramx.Y = 30

Call LoadAntiCheat

End Sub

Sub WriteVar(ByVal File As String, ByVal Main As String, ByVal Var As String, ByVal Value As String)
'*****************************************************************
'Escribe VAR en un archivo
'*****************************************************************

writeprivateprofilestring Main, Var, Value, File
    
End Sub

Sub SaveUser(ByVal UserIndex As Integer, ByVal UserFile As String)
'*************************************************
'Author: Unknown
'Last modified: 23/01/2007
'Saves the Users records
'23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
'*************************************************

On Error GoTo Errhandler

Dim OldUserHead As Long


'ESTO TIENE QUE EVITAR ESE BUGAZO QUE NO SE POR QUE GRABA USUARIOS NULOS
'clase=0 es el error, porq el enum empieza de 1!!
If UserList(UserIndex).clase = 0 Or UserList(UserIndex).Stats.ELV = 0 Then
    Call LogCriticEvent("Estoy intentantdo guardar un usuario nulo de nombre: " & UserList(UserIndex).Name)
    Exit Sub
End If


If UserList(UserIndex).flags.Mimetizado = 1 Then
    UserList(UserIndex).Char.body = UserList(UserIndex).CharMimetizado.body
    UserList(UserIndex).Char.Head = UserList(UserIndex).CharMimetizado.Head
    UserList(UserIndex).Char.CascoAnim = UserList(UserIndex).CharMimetizado.CascoAnim
    UserList(UserIndex).Char.ShieldAnim = UserList(UserIndex).CharMimetizado.ShieldAnim
    UserList(UserIndex).Char.WeaponAnim = UserList(UserIndex).CharMimetizado.WeaponAnim
    UserList(UserIndex).Counters.Mimetismo = 0
    UserList(UserIndex).flags.Mimetizado = 0
End If



If FileExist(UserFile, vbNormal) Then
    If UserList(UserIndex).flags.Muerto = 1 Then
        OldUserHead = UserList(UserIndex).Char.Head
        UserList(UserIndex).Char.Head = GetVar(UserFile, "INIT", "Head")
    End If
'       Kill UserFile
End If

Dim LoopC As Integer


Call WriteVar(UserFile, "FLAGS", "Muerto", CStr(UserList(UserIndex).flags.Muerto))
Call WriteVar(UserFile, "FLAGS", "Escondido", CStr(UserList(UserIndex).flags.Escondido))
Call WriteVar(UserFile, "FLAGS", "Hambre", CStr(UserList(UserIndex).flags.Hambre))
Call WriteVar(UserFile, "FLAGS", "Sed", CStr(UserList(UserIndex).flags.Sed))
Call WriteVar(UserFile, "FLAGS", "Desnudo", CStr(UserList(UserIndex).flags.Desnudo))
Call WriteVar(UserFile, "FLAGS", "Ban", CStr(UserList(UserIndex).flags.Ban))

Call WriteVar(UserFile, "FLAGS", "Navegando", CStr(UserList(UserIndex).flags.Navegando))
Call WriteVar(UserFile, "FLAGS", "Equitando", CStr(UserList(UserIndex).flags.Equitando))

Call WriteVar(UserFile, "FLAGS", "Metamorfosis", CStr(UserList(UserIndex).flags.Metamorfosis)) 'Metamorfosis

Call WriteVar(UserFile, "FLAGS", "Envenenado", CStr(UserList(UserIndex).flags.Envenenado))
Call WriteVar(UserFile, "FLAGS", "Paralizado", CStr(UserList(UserIndex).flags.Paralizado))

Call WriteVar(UserFile, "CONSEJO", "PERTENECE", IIf(UserList(UserIndex).flags.Privilegios And PlayerType.RoyalCouncil, "1", "0"))
Call WriteVar(UserFile, "CONSEJO", "PERTENECECAOS", IIf(UserList(UserIndex).flags.Privilegios And PlayerType.ChaosCouncil, "1", "0"))


Call WriteVar(UserFile, "COUNTERS", "Pena", CStr(UserList(UserIndex).Counters.Pena))
Call WriteVar(UserFile, "COUNTERS", "Piedra", CStr(UserList(UserIndex).Counters.Piedra))

Call WriteVar(UserFile, "FACCIONES", "EjercitoReal", CStr(UserList(UserIndex).Faccion.ArmadaReal))
Call WriteVar(UserFile, "FACCIONES", "EjercitoCaos", CStr(UserList(UserIndex).Faccion.FuerzasCaos))
Call WriteVar(UserFile, "FACCIONES", "CiudMatados", CStr(UserList(UserIndex).Faccion.CiudadanosMatados))
Call WriteVar(UserFile, "FACCIONES", "CrimMatados", CStr(UserList(UserIndex).Faccion.CriminalesMatados))
Call WriteVar(UserFile, "FACCIONES", "rArCaos", CStr(UserList(UserIndex).Faccion.RecibioArmaduraCaos))
Call WriteVar(UserFile, "FACCIONES", "rArReal", CStr(UserList(UserIndex).Faccion.RecibioArmaduraReal))
Call WriteVar(UserFile, "FACCIONES", "rExCaos", CStr(UserList(UserIndex).Faccion.RecibioExpInicialCaos))
Call WriteVar(UserFile, "FACCIONES", "rExReal", CStr(UserList(UserIndex).Faccion.RecibioExpInicialReal))
Call WriteVar(UserFile, "FACCIONES", "recCaos", CStr(UserList(UserIndex).Faccion.RecompensasCaos))
Call WriteVar(UserFile, "FACCIONES", "recReal", CStr(UserList(UserIndex).Faccion.RecompensasReal))
Call WriteVar(UserFile, "FACCIONES", "Reenlistadas", CStr(UserList(UserIndex).Faccion.Reenlistadas))
Call WriteVar(UserFile, "FACCIONES", "NivelIngreso", CStr(UserList(UserIndex).Faccion.NivelIngreso))
Call WriteVar(UserFile, "FACCIONES", "FechaIngreso", UserList(UserIndex).Faccion.FechaIngreso)
Call WriteVar(UserFile, "FACCIONES", "MatadosIngreso", CStr(UserList(UserIndex).Faccion.MatadosIngreso))
Call WriteVar(UserFile, "FACCIONES", "NextRecompensa", CStr(UserList(UserIndex).Faccion.NextRecompensa))


'¿Fueron modificados los atributos del usuario?
If Not UserList(UserIndex).flags.TomoPocion Then
    For LoopC = 1 To UBound(UserList(UserIndex).Stats.UserAtributos)
        Call WriteVar(UserFile, "ATRIBUTOS", "AT" & LoopC, CStr(UserList(UserIndex).Stats.UserAtributos(LoopC)))
    Next
Else
    For LoopC = 1 To UBound(UserList(UserIndex).Stats.UserAtributos)
        'UserList(UserIndex).Stats.UserAtributos(LoopC) = UserList(UserIndex).Stats.UserAtributosBackUP(LoopC)
        Call WriteVar(UserFile, "ATRIBUTOS", "AT" & LoopC, CStr(UserList(UserIndex).Stats.UserAtributosBackUP(LoopC)))
    Next
End If

For LoopC = 1 To UBound(UserList(UserIndex).Stats.UserSkills)
    Call WriteVar(UserFile, "SKILLS", "SK" & LoopC, CStr(UserList(UserIndex).Stats.UserSkills(LoopC)))
Next


Call WriteVar(UserFile, "CONTACTO", "Email", UserList(UserIndex).email)

Call WriteVar(UserFile, "INIT", "Genero", UserList(UserIndex).genero)
Call WriteVar(UserFile, "INIT", "Raza", UserList(UserIndex).raza)
Call WriteVar(UserFile, "INIT", "Hogar", UserList(UserIndex).Hogar)
Call WriteVar(UserFile, "INIT", "Clase", UserList(UserIndex).clase)
Call WriteVar(UserFile, "INIT", "Desc", UserList(UserIndex).desc)

Call WriteVar(UserFile, "INIT", "Heading", CStr(UserList(UserIndex).Char.heading))

Call WriteVar(UserFile, "INIT", "Head", CStr(UserList(UserIndex).OrigChar.Head))

If UserList(UserIndex).flags.Muerto = 0 Then
    Call WriteVar(UserFile, "INIT", "Body", CStr(UserList(UserIndex).Char.body))
End If

Call WriteVar(UserFile, "INIT", "Arma", CStr(UserList(UserIndex).Char.WeaponAnim))
Call WriteVar(UserFile, "INIT", "Escudo", CStr(UserList(UserIndex).Char.ShieldAnim))
Call WriteVar(UserFile, "INIT", "Casco", CStr(UserList(UserIndex).Char.CascoAnim))
Call WriteVar(UserFile, "INIT", "Aura", CStr(UserList(UserIndex).Char.Aura))

#If ConUpTime Then
    Dim TempDate As Date
    TempDate = Now - UserList(UserIndex).LogOnTime
    UserList(UserIndex).LogOnTime = Now
    UserList(UserIndex).UpTime = UserList(UserIndex).UpTime + (Abs(Day(TempDate) - 30) * 24 * 3600) + Hour(TempDate) * 3600 + Minute(TempDate) * 60 + Second(TempDate)
    UserList(UserIndex).UpTime = UserList(UserIndex).UpTime
    Call WriteVar(UserFile, "INIT", "UpTime", UserList(UserIndex).UpTime)
#End If

'First time around?
If GetVar(UserFile, "INIT", "LastIP1") = vbNullString Then
    Call WriteVar(UserFile, "INIT", "LastIP1", UserList(UserIndex).ip & " - " & Date & ":" & Time)
'Is it a different ip from last time?
ElseIf UserList(UserIndex).ip <> Left$(GetVar(UserFile, "INIT", "LastIP1"), InStr(1, GetVar(UserFile, "INIT", "LastIP1"), " ") - 1) Then
    Dim i As Integer
    For i = 5 To 2 Step -1
        Call WriteVar(UserFile, "INIT", "LastIP" & i, GetVar(UserFile, "INIT", "LastIP" & CStr(i - 1)))
    Next i
    Call WriteVar(UserFile, "INIT", "LastIP1", UserList(UserIndex).ip & " - " & Date & ":" & Time)
'Same ip, just update the date
Else
    Call WriteVar(UserFile, "INIT", "LastIP1", UserList(UserIndex).ip & " - " & Date & ":" & Time)
End If



Call WriteVar(UserFile, "INIT", "Position", UserList(UserIndex).Pos.map & "-" & UserList(UserIndex).Pos.X & "-" & UserList(UserIndex).Pos.Y)


Call WriteVar(UserFile, "STATS", "GLD", CStr(UserList(UserIndex).Stats.GLD))
Call WriteVar(UserFile, "STATS", "Puntos", CStr(UserList(UserIndex).Stats.Puntos))

Call WriteVar(UserFile, "STATS", "BANCO", CStr(UserList(UserIndex).Stats.Banco))

Call WriteVar(UserFile, "STATS", "MaxHP", CStr(UserList(UserIndex).Stats.MaxHP))
Call WriteVar(UserFile, "STATS", "MinHP", CStr(UserList(UserIndex).Stats.MinHP))

Call WriteVar(UserFile, "STATS", "MaxSTA", CStr(UserList(UserIndex).Stats.MaxSta))
Call WriteVar(UserFile, "STATS", "MinSTA", CStr(UserList(UserIndex).Stats.MinSta))

Call WriteVar(UserFile, "STATS", "MaxMAN", CStr(UserList(UserIndex).Stats.MaxMAN))
Call WriteVar(UserFile, "STATS", "MinMAN", CStr(UserList(UserIndex).Stats.MinMAN))

Call WriteVar(UserFile, "STATS", "MaxHIT", CStr(UserList(UserIndex).Stats.MaxHIT))
Call WriteVar(UserFile, "STATS", "MinHIT", CStr(UserList(UserIndex).Stats.MinHIT))

Call WriteVar(UserFile, "STATS", "MaxAGU", CStr(UserList(UserIndex).Stats.MaxAGU))
Call WriteVar(UserFile, "STATS", "MinAGU", CStr(UserList(UserIndex).Stats.MinAGU))

Call WriteVar(UserFile, "STATS", "MaxHAM", CStr(UserList(UserIndex).Stats.MaxHam))
Call WriteVar(UserFile, "STATS", "MinHAM", CStr(UserList(UserIndex).Stats.MinHam))

Call WriteVar(UserFile, "STATS", "SkillPtsLibres", CStr(UserList(UserIndex).Stats.SkillPts))
  
Call WriteVar(UserFile, "STATS", "EXP", CStr(UserList(UserIndex).Stats.Exp))
Call WriteVar(UserFile, "STATS", "ELV", CStr(UserList(UserIndex).Stats.ELV))





Call WriteVar(UserFile, "STATS", "ELU", CStr(UserList(UserIndex).Stats.ELU))
Call WriteVar(UserFile, "MUERTES", "UserMuertes", CStr(UserList(UserIndex).Stats.UsuariosMatados))
'Call WriteVar(UserFile, "MUERTES", "CrimMuertes", CStr(UserList(UserIndex).Stats.CriminalesMatados))
Call WriteVar(UserFile, "MUERTES", "NpcsMuertes", CStr(UserList(UserIndex).Stats.NPCsMuertos))
  
'[KEVIN]----------------------------------------------------------------------------
'*******************************************************************************************
Call WriteVar(UserFile, "BancoInventory", "CantidadItems", val(UserList(UserIndex).BancoInvent.NroItems))
Dim loopd As Integer
For loopd = 1 To MAX_BANCOINVENTORY_SLOTS
    Call WriteVar(UserFile, "BancoInventory", "Obj" & loopd, UserList(UserIndex).BancoInvent.Object(loopd).ObjIndex & "-" & UserList(UserIndex).BancoInvent.Object(loopd).amount)
Next loopd
'*******************************************************************************************
'[/KEVIN]-----------
  
'Save Inv
Call WriteVar(UserFile, "Inventory", "CantidadItems", val(UserList(UserIndex).Invent.NroItems))

For LoopC = 1 To MAX_INVENTORY_SLOTS
    Call WriteVar(UserFile, "Inventory", "Obj" & LoopC, UserList(UserIndex).Invent.Object(LoopC).ObjIndex & "-" & UserList(UserIndex).Invent.Object(LoopC).amount & "-" & UserList(UserIndex).Invent.Object(LoopC).Equipped)
Next

Call WriteVar(UserFile, "Inventory", "WeaponEqpSlot", CStr(UserList(UserIndex).Invent.WeaponEqpSlot))
Call WriteVar(UserFile, "Inventory", "ArmourEqpSlot", CStr(UserList(UserIndex).Invent.ArmourEqpSlot))
Call WriteVar(UserFile, "Inventory", "CascoEqpSlot", CStr(UserList(UserIndex).Invent.CascoEqpSlot))
Call WriteVar(UserFile, "Inventory", "EscudoEqpSlot", CStr(UserList(UserIndex).Invent.EscudoEqpSlot))
Call WriteVar(UserFile, "Inventory", "BarcoSlot", CStr(UserList(UserIndex).Invent.BarcoSlot))
Call WriteVar(UserFile, "Inventory", "MonturaSlot", CStr(UserList(UserIndex).Invent.MonturaSlot))
Call WriteVar(UserFile, "Inventory", "MunicionSlot", CStr(UserList(UserIndex).Invent.MunicionEqpSlot))
'/Nacho

Call WriteVar(UserFile, "Inventory", "AnilloSlot", CStr(UserList(UserIndex).Invent.AnilloEqpSlot))


'Reputacion
Call WriteVar(UserFile, "REP", "Asesino", CStr(UserList(UserIndex).Reputacion.AsesinoRep))
Call WriteVar(UserFile, "REP", "Bandido", CStr(UserList(UserIndex).Reputacion.BandidoRep))
Call WriteVar(UserFile, "REP", "Burguesia", CStr(UserList(UserIndex).Reputacion.BurguesRep))
Call WriteVar(UserFile, "REP", "Ladrones", CStr(UserList(UserIndex).Reputacion.LadronesRep))
Call WriteVar(UserFile, "REP", "Nobles", CStr(UserList(UserIndex).Reputacion.NobleRep))
Call WriteVar(UserFile, "REP", "Plebe", CStr(UserList(UserIndex).Reputacion.PlebeRep))

Dim L As Long
L = (-UserList(UserIndex).Reputacion.AsesinoRep) + _
    (-UserList(UserIndex).Reputacion.BandidoRep) + _
    UserList(UserIndex).Reputacion.BurguesRep + _
    (-UserList(UserIndex).Reputacion.LadronesRep) + _
    UserList(UserIndex).Reputacion.NobleRep + _
    UserList(UserIndex).Reputacion.PlebeRep
L = L / 6
Call WriteVar(UserFile, "REP", "Promedio", CStr(L))

Dim cad As String

For LoopC = 1 To MAXUSERHECHIZOS
    cad = UserList(UserIndex).Stats.UserHechizos(LoopC)
    Call WriteVar(UserFile, "HECHIZOS", "H" & LoopC, cad)
Next

Call SaveQuestStats(UserIndex, UserFile)

Dim NroMascotas As Long
NroMascotas = UserList(UserIndex).NroMascotas

For LoopC = 1 To MAXMASCOTAS
    ' Mascota valida?
    If UserList(UserIndex).MascotasIndex(LoopC) > 0 Then
        ' Nos aseguramos que la criatura no fue invocada
        If Npclist(UserList(UserIndex).MascotasIndex(LoopC)).Contadores.TiempoExistencia = 0 Then
            cad = UserList(UserIndex).MascotasType(LoopC)
        Else 'Si fue invocada no la guardamos
            cad = "0"
            NroMascotas = NroMascotas - 1
        End If
        Call WriteVar(UserFile, "MASCOTAS", "MAS" & LoopC, cad)
    Else
        cad = UserList(UserIndex).MascotasType(LoopC)
        Call WriteVar(UserFile, "MASCOTAS", "MAS" & LoopC, cad)
    End If

Next

Call WriteVar(UserFile, "MASCOTAS", "NroMascotas", CStr(NroMascotas))

'Devuelve el head de muerto
If UserList(UserIndex).flags.Muerto = 1 Then
    UserList(UserIndex).Char.Head = iCabezaMuerto
End If

Exit Sub

Errhandler:
Call LogError("Error en SaveUser")

End Sub

Function criminal(ByVal UserIndex As Integer) As Boolean

Dim L As Long
L = (-UserList(UserIndex).Reputacion.AsesinoRep) + _
    (-UserList(UserIndex).Reputacion.BandidoRep) + _
    UserList(UserIndex).Reputacion.BurguesRep + _
    (-UserList(UserIndex).Reputacion.LadronesRep) + _
    UserList(UserIndex).Reputacion.NobleRep + _
    UserList(UserIndex).Reputacion.PlebeRep
L = L / 6
criminal = (L < 0)

End Function

Sub BackUPnPc(NpcIndex As Integer)

Dim NpcNumero As Integer
Dim npcfile As String
Dim LoopC As Integer


NpcNumero = Npclist(NpcIndex).Numero

'If NpcNumero > 499 Then
'    npcfile = DatPath & "bkNPCs-HOSTILES.dat"
'Else
    npcfile = DatPath & "bkNPCs.dat"
'End If

'General
Call WriteVar(npcfile, "NPC" & NpcNumero, "Name", Npclist(NpcIndex).Name)
Call WriteVar(npcfile, "NPC" & NpcNumero, "Desc", Npclist(NpcIndex).desc)
Call WriteVar(npcfile, "NPC" & NpcNumero, "Head", val(Npclist(NpcIndex).Char.Head))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Body", val(Npclist(NpcIndex).Char.body))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Heading", val(Npclist(NpcIndex).Char.heading))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Movement", val(Npclist(NpcIndex).Movement))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Attackable", val(Npclist(NpcIndex).Attackable))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Comercia", val(Npclist(NpcIndex).Comercia))
Call WriteVar(npcfile, "NPC" & NpcNumero, "TipoItems", val(Npclist(NpcIndex).TipoItems))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Hostil", val(Npclist(NpcIndex).Hostile))
Call WriteVar(npcfile, "NPC" & NpcNumero, "GiveEXP", val(Npclist(NpcIndex).GiveEXP))
Call WriteVar(npcfile, "NPC" & NpcNumero, "GiveGLD", val(Npclist(NpcIndex).GiveGLD))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Hostil", val(Npclist(NpcIndex).Hostile))
Call WriteVar(npcfile, "NPC" & NpcNumero, "InvReSpawn", val(Npclist(NpcIndex).InvReSpawn))
Call WriteVar(npcfile, "NPC" & NpcNumero, "NpcType", val(Npclist(NpcIndex).NPCtype))


'Stats
Call WriteVar(npcfile, "NPC" & NpcNumero, "Alineacion", val(Npclist(NpcIndex).Stats.Alineacion))
Call WriteVar(npcfile, "NPC" & NpcNumero, "DEF", val(Npclist(NpcIndex).Stats.def))
Call WriteVar(npcfile, "NPC" & NpcNumero, "MaxHit", val(Npclist(NpcIndex).Stats.MaxHIT))
Call WriteVar(npcfile, "NPC" & NpcNumero, "MaxHp", val(Npclist(NpcIndex).Stats.MaxHP))
Call WriteVar(npcfile, "NPC" & NpcNumero, "MinHit", val(Npclist(NpcIndex).Stats.MinHIT))
Call WriteVar(npcfile, "NPC" & NpcNumero, "MinHp", val(Npclist(NpcIndex).Stats.MinHP))




'Flags
Call WriteVar(npcfile, "NPC" & NpcNumero, "ReSpawn", val(Npclist(NpcIndex).flags.Respawn))
Call WriteVar(npcfile, "NPC" & NpcNumero, "BackUp", val(Npclist(NpcIndex).flags.BackUp))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Domable", val(Npclist(NpcIndex).flags.Domable))

'Inventario
Call WriteVar(npcfile, "NPC" & NpcNumero, "NroItems", val(Npclist(NpcIndex).Invent.NroItems))
If Npclist(NpcIndex).Invent.NroItems > 0 Then
   For LoopC = 1 To MAX_INVENTORY_SLOTS
        Call WriteVar(npcfile, "NPC" & NpcNumero, "Obj" & LoopC, Npclist(NpcIndex).Invent.Object(LoopC).ObjIndex & "-" & Npclist(NpcIndex).Invent.Object(LoopC).amount)
   Next
End If


End Sub



Sub CargarNpcBackUp(NpcIndex As Integer, ByVal NpcNumber As Integer)

'Status
If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando backup Npc"

Dim npcfile As String

'If NpcNumber > 499 Then
'    npcfile = DatPath & "bkNPCs-HOSTILES.dat"
'Else
    npcfile = DatPath & "bkNPCs.dat"
'End If

Npclist(NpcIndex).Numero = NpcNumber
Npclist(NpcIndex).Name = GetVar(npcfile, "NPC" & NpcNumber, "Name")
Npclist(NpcIndex).desc = GetVar(npcfile, "NPC" & NpcNumber, "Desc")
Npclist(NpcIndex).Movement = val(GetVar(npcfile, "NPC" & NpcNumber, "Movement"))
Npclist(NpcIndex).NPCtype = val(GetVar(npcfile, "NPC" & NpcNumber, "NpcType"))

Npclist(NpcIndex).Char.body = val(GetVar(npcfile, "NPC" & NpcNumber, "Body"))
Npclist(NpcIndex).Char.Head = val(GetVar(npcfile, "NPC" & NpcNumber, "Head"))
Npclist(NpcIndex).Char.heading = val(GetVar(npcfile, "NPC" & NpcNumber, "Heading"))

Npclist(NpcIndex).Attackable = val(GetVar(npcfile, "NPC" & NpcNumber, "Attackable"))
Npclist(NpcIndex).Comercia = val(GetVar(npcfile, "NPC" & NpcNumber, "Comercia"))
Npclist(NpcIndex).Hostile = val(GetVar(npcfile, "NPC" & NpcNumber, "Hostile"))
Npclist(NpcIndex).GiveEXP = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveEXP")) * ExpX


Npclist(NpcIndex).GiveGLD = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveGLD")) * OroX

Npclist(NpcIndex).InvReSpawn = val(GetVar(npcfile, "NPC" & NpcNumber, "InvReSpawn"))

Npclist(NpcIndex).Stats.MaxHP = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHP"))
Npclist(NpcIndex).Stats.MinHP = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHP"))
Npclist(NpcIndex).Stats.MaxHIT = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHIT"))
Npclist(NpcIndex).Stats.MinHIT = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHIT"))
Npclist(NpcIndex).Stats.def = val(GetVar(npcfile, "NPC" & NpcNumber, "DEF"))
Npclist(NpcIndex).Stats.Alineacion = val(GetVar(npcfile, "NPC" & NpcNumber, "Alineacion"))



Dim LoopC As Integer
Dim ln As String
Npclist(NpcIndex).Invent.NroItems = val(GetVar(npcfile, "NPC" & NpcNumber, "NROITEMS"))
If Npclist(NpcIndex).Invent.NroItems > 0 Then
    For LoopC = 1 To MAX_INVENTORY_SLOTS
        ln = GetVar(npcfile, "NPC" & NpcNumber, "Obj" & LoopC)
        Npclist(NpcIndex).Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
        Npclist(NpcIndex).Invent.Object(LoopC).amount = val(ReadField(2, ln, 45))
       
    Next LoopC
Else
    For LoopC = 1 To MAX_INVENTORY_SLOTS
        Npclist(NpcIndex).Invent.Object(LoopC).ObjIndex = 0
        Npclist(NpcIndex).Invent.Object(LoopC).amount = 0
    Next LoopC
End If



Npclist(NpcIndex).flags.NPCActive = True
Npclist(NpcIndex).flags.Respawn = val(GetVar(npcfile, "NPC" & NpcNumber, "ReSpawn"))
Npclist(NpcIndex).flags.BackUp = val(GetVar(npcfile, "NPC" & NpcNumber, "BackUp"))
Npclist(NpcIndex).flags.Domable = val(GetVar(npcfile, "NPC" & NpcNumber, "Domable"))
Npclist(NpcIndex).flags.RespawnOrigPos = val(GetVar(npcfile, "NPC" & NpcNumber, "OrigPos"))

'Tipo de items con los que comercia
Npclist(NpcIndex).TipoItems = val(GetVar(npcfile, "NPC" & NpcNumber, "TipoItems"))

End Sub


Sub LogBan(ByVal BannedIndex As Integer, ByVal UserIndex As Integer, ByVal Motivo As String)

Call WriteVar(App.Path & "\logs\" & "BanDetail.log", UserList(BannedIndex).Name, "BannedBy", UserList(UserIndex).Name)
Call WriteVar(App.Path & "\logs\" & "BanDetail.log", UserList(BannedIndex).Name, "Reason", Motivo)

'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
Dim mifile As Integer
mifile = FreeFile
Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
Print #mifile, UserList(BannedIndex).Name
Close #mifile

End Sub


Sub LogBanFromName(ByVal BannedName As String, ByVal UserIndex As Integer, ByVal Motivo As String)

Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "BannedBy", UserList(UserIndex).Name)
Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "Reason", Motivo)

'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
Dim mifile As Integer
mifile = FreeFile
Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
Print #mifile, BannedName
Close #mifile

End Sub


Sub Ban(ByVal BannedName As String, ByVal Baneador As String, ByVal Motivo As String)

Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "BannedBy", Baneador)
Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "Reason", Motivo)


'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
Dim mifile As Integer
mifile = FreeFile
Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
Print #mifile, BannedName
Close #mifile

End Sub

Public Sub CargaApuestas()

    Apuestas.Ganancias = val(GetVar(DatPath & "apuestas.dat", "Main", "Ganancias"))
    Apuestas.Perdidas = val(GetVar(DatPath & "apuestas.dat", "Main", "Perdidas"))
    Apuestas.Jugadas = val(GetVar(DatPath & "apuestas.dat", "Main", "Jugadas"))

End Sub
