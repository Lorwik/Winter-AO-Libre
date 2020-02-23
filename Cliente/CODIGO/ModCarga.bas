Attribute VB_Name = "ModCarga"
'*******************************************MODULO DE CARGA*********************************************
'AUTOR: MANUEL (LORWIK)
'DESCRIPCION: RECOPILACION DE TODOS LOS CODIGOS NECESARIOS PARA LA CARGA DEL CLIENTE
'*******************************************************************************************************

Option Explicit

Public Type tCabecera 'Cabecera de los con
    desc As String * 255
    CRC As Long
    MagicWord As Long
End Type

Public Type tSetupMods
    byMemory    As Byte
    bUseVideo   As Boolean
    bNoRes      As Boolean
End Type

Public ClientSetup As tSetupMods

Public MiCabecera As tCabecera
Private file As String

Sub CargarCabezas()
    Dim N As Integer
    Dim i As Long
    Dim j As Byte
    Dim Numheads As Integer
    Dim Miscabezas() As tIndiceCabeza
        
    'Cabezas
    file = Get_Extract(Scripts, "Cabezas.ind")
    N = FreeFile()
    Open file For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , Numheads
    
    'Resize array
    ReDim HeadData(0 To Numheads) As HeadData
    ReDim Miscabezas(0 To Numheads) As tIndiceCabeza
    
    For i = 1 To Numheads
        Get #N, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            For j = 1 To 4
                Call InitGrh(HeadData(i).Head(j), Miscabezas(i).Head(j), 0)
            Next j
        End If
    Next i
    
    Close #N
    
    Delete_File file
End Sub

Sub CargarCascos()
    Dim N As Integer
    Dim i As Long
    Dim j As Byte
    Dim NumCascos As Integer

    Dim Miscabezas() As tIndiceCabeza
    
    file = Get_Extract(Scripts, "Cascos.ind")
    N = FreeFile()
    Open file For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCascos
    
    'Resize array
    ReDim CascoAnimData(0 To NumCascos) As HeadData
    ReDim Miscabezas(0 To NumCascos) As tIndiceCabeza
    
    For i = 1 To NumCascos
        Get #N, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            For j = 1 To 4
                Call InitGrh(CascoAnimData(i).Head(j), Miscabezas(i).Head(j), 0)
            Next j
        End If
    Next i
    
    Close #N
    
    Delete_File file
End Sub

Sub CargarCuerpos()
    Dim N As Integer
    Dim i As Long
    Dim j As Byte
    Dim NumCuerpos As Integer
    Dim MisCuerpos() As tIndiceCuerpo
    
    N = FreeFile()
    file = Get_Extract(Scripts, "Personajes.ind")
    Open file For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCuerpos
    
    'Resize array
    ReDim BodyData(0 To NumCuerpos) As BodyData
    ReDim MisCuerpos(0 To NumCuerpos) As tIndiceCuerpo
    
    For i = 1 To NumCuerpos
        Get #N, , MisCuerpos(i)
        
        If MisCuerpos(i).Body(1) Then
            For j = 1 To 4
                InitGrh BodyData(i).Walk(j), MisCuerpos(i).Body(j), 0
            Next j
            BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
            BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY
        End If
    Next i
    
    Close #N
    
    Delete_File file
End Sub

Sub CargarFxs()
    Dim N As Integer
    Dim i As Long
    Dim NumFxs As Integer
    
    file = Get_Extract(Scripts, "Fxs.ind")
    
    N = FreeFile()
    Open file For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumFxs
    
    'Resize array
    ReDim FxData(1 To NumFxs) As tIndiceFx
    
    For i = 1 To NumFxs
        Get #N, , FxData(i)
    Next i
    
    Close #N
    
    Delete_File file
End Sub

Sub CargarAnimArmas()
    On Error Resume Next
    Dim i As Integer
    Dim Archivo As String
    Dim Leer As New clsIniReader
    
    'Armas
    file = Get_Extract(Scripts, "armas.dat")
    Archivo = file
    Leer.Initialize Archivo
    NumWeaponAnims = Val(Leer.GetValue("INIT", "NumArmas"))
            
    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
    For i = 1 To NumWeaponAnims
        InitGrh WeaponAnimData(i).WeaponWalk(1), Val(Leer.GetValue("ARMA" & i, "Dir1")), 0
        InitGrh WeaponAnimData(i).WeaponWalk(2), Val(Leer.GetValue("ARMA" & i, "Dir2")), 0
        InitGrh WeaponAnimData(i).WeaponWalk(3), Val(Leer.GetValue("ARMA" & i, "Dir3")), 0
        InitGrh WeaponAnimData(i).WeaponWalk(4), Val(Leer.GetValue("ARMA" & i, "Dir4")), 0
    Next i
        Delete_File file

End Sub


Sub CargarColores()
On Error Resume Next
    Dim ArchivoC As String
    Dim Leer As New clsIniReader
    
    file = Get_Extract(Scripts, "colores.dat")
    ArchivoC = file
    Leer.Initialize ArchivoC
    
    If Not FileExist(ArchivoC, vbArchive) Then
'TODO : Si hay que reinstalar, porque no cierra???
        Call MsgBox("ERROR: no se ha podido cargar los colores. Falta el archivo colores.dat, reinstale el juego", vbCritical + vbOKOnly)
        Exit Sub
    End If
    
    Dim i As Long
    
    For i = 0 To 48 '49 y 50 reservados para ciudadano y criminal
        ColoresPJ(i).r = CByte(GetVar(ArchivoC, CStr(i), "R"))
        ColoresPJ(i).g = CByte(GetVar(ArchivoC, CStr(i), "G"))
        ColoresPJ(i).B = CByte(GetVar(ArchivoC, CStr(i), "B"))
    Next i
    
    ColoresPJ(50).r = CByte(GetVar(ArchivoC, "CR", "R"))
    ColoresPJ(50).g = CByte(GetVar(ArchivoC, "CR", "G"))
    ColoresPJ(50).B = CByte(GetVar(ArchivoC, "CR", "B"))
    ColoresPJ(49).r = CByte(GetVar(ArchivoC, "CI", "R"))
    ColoresPJ(49).g = CByte(GetVar(ArchivoC, "CI", "G"))
    ColoresPJ(49).B = CByte(GetVar(ArchivoC, "CI", "B"))
End Sub

Sub CargarAnimEscudos()
    Dim i As Integer
    Dim Archivo As String
    Dim Leer As New clsIniReader
    
    'Escudos
    file = Get_Extract(Scripts, "escudos.dat")
    Archivo = file
    
    Leer.Initialize Archivo
    
    NumEscudosAnims = Val(Leer.GetValue("INIT", "NumEscudos"))
            
    ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
    For i = 1 To NumEscudosAnims
        InitGrh ShieldAnimData(i).ShieldWalk(1), Val(Leer.GetValue("ESC" & i, "Dir1")), 0
        InitGrh ShieldAnimData(i).ShieldWalk(2), Val(Leer.GetValue("ESC" & i, "Dir2")), 0
        InitGrh ShieldAnimData(i).ShieldWalk(3), Val(Leer.GetValue("ESC" & i, "Dir3")), 0
        InitGrh ShieldAnimData(i).ShieldWalk(4), Val(Leer.GetValue("ESC" & i, "Dir4")), 0
    Next i
    Delete_File file
    
End Sub
Public Sub InicializarNombres()

    ListaRazas(eRaza.Humano) = "Humano"
    ListaRazas(eRaza.Elfo) = "Elfo"
    ListaRazas(eRaza.ElfoOscuro) = "Elfo Oscuro"
    ListaRazas(eRaza.Gnomo) = "Gnomo"
    ListaRazas(eRaza.Enano) = "Enano"
    ListaRazas(eRaza.Orco) = "Orco"
    
    ListaClases(eClass.Mage) = "Mago"
    ListaClases(eClass.Cleric) = "Clerigo"
    ListaClases(eClass.Warrior) = "Guerrero"
    ListaClases(eClass.Assasin) = "Asesino"
    ListaClases(eClass.Thief) = "Ladron"
    ListaClases(eClass.Bard) = "Bardo"
    ListaClases(eClass.Druid) = "Druida"
    ListaClases(eClass.Bandit) = "Bandido"
    ListaClases(eClass.Paladin) = "Paladin"
    ListaClases(eClass.Hunter) = "Cazador"
    ListaClases(eClass.Trabajador) = "Trabajador"
    ListaClases(eClass.Pirat) = "Pirata"
End Sub

Sub CargarParticulas()
'*********************************
'Carga de particulas.
'*********************************
    Dim StreamFile As String
    Dim loopc As Long
    Dim i As Long
    Dim GrhListing As String
    Dim TempSet As String
    Dim ColorSet As Long
    Dim Leer As New clsIniReader

    file = Get_Extract(Scripts, "particles.ini")
    
    StreamFile = file
    
    Leer.Initialize StreamFile

    TotalStreams = Val(Leer.GetValue("INIT", "Total"))
     
    'resize StreamData array
    ReDim StreamData(1 To TotalStreams) As Stream
     
        'fill StreamData array with info from Particles.ini
        For loopc = 1 To TotalStreams
            StreamData(loopc).name = General_Var_Get(StreamFile, Val(loopc), "Name")
            StreamData(loopc).NumOfParticles = General_Var_Get(StreamFile, Val(loopc), "NumOfParticles")
            StreamData(loopc).x1 = General_Var_Get(StreamFile, Val(loopc), "X1")
            StreamData(loopc).y1 = General_Var_Get(StreamFile, Val(loopc), "Y1")
            StreamData(loopc).x2 = General_Var_Get(StreamFile, Val(loopc), "X2")
            StreamData(loopc).y2 = General_Var_Get(StreamFile, Val(loopc), "Y2")
            StreamData(loopc).angle = General_Var_Get(StreamFile, Val(loopc), "Angle")
            StreamData(loopc).vecx1 = General_Var_Get(StreamFile, Val(loopc), "VecX1")
            StreamData(loopc).vecx2 = General_Var_Get(StreamFile, Val(loopc), "VecX2")
            StreamData(loopc).vecy1 = General_Var_Get(StreamFile, Val(loopc), "VecY1")
            StreamData(loopc).vecy2 = General_Var_Get(StreamFile, Val(loopc), "VecY2")
            StreamData(loopc).life1 = General_Var_Get(StreamFile, Val(loopc), "Life1")
            StreamData(loopc).life2 = General_Var_Get(StreamFile, Val(loopc), "Life2")
            StreamData(loopc).friction = General_Var_Get(StreamFile, Val(loopc), "Friction")
            StreamData(loopc).spin = General_Var_Get(StreamFile, Val(loopc), "Spin")
            StreamData(loopc).spin_speedL = General_Var_Get(StreamFile, Val(loopc), "Spin_SpeedL")
            StreamData(loopc).spin_speedH = General_Var_Get(StreamFile, Val(loopc), "Spin_SpeedH")
            StreamData(loopc).AlphaBlend = General_Var_Get(StreamFile, Val(loopc), "AlphaBlend")
            StreamData(loopc).gravity = General_Var_Get(StreamFile, Val(loopc), "Gravity")
            StreamData(loopc).grav_strength = General_Var_Get(StreamFile, Val(loopc), "Grav_Strength")
            StreamData(loopc).bounce_strength = General_Var_Get(StreamFile, Val(loopc), "Bounce_Strength")
            StreamData(loopc).XMove = General_Var_Get(StreamFile, Val(loopc), "XMove")
            StreamData(loopc).YMove = General_Var_Get(StreamFile, Val(loopc), "YMove")
            StreamData(loopc).move_x1 = General_Var_Get(StreamFile, Val(loopc), "move_x1")
            StreamData(loopc).move_x2 = General_Var_Get(StreamFile, Val(loopc), "move_x2")
            StreamData(loopc).move_y1 = General_Var_Get(StreamFile, Val(loopc), "move_y1")
            StreamData(loopc).move_y2 = General_Var_Get(StreamFile, Val(loopc), "move_y2")
            StreamData(loopc).Radio = Val(General_Var_Get(StreamFile, Val(loopc), "Radio"))
            StreamData(loopc).life_counter = General_Var_Get(StreamFile, Val(loopc), "life_counter")
            StreamData(loopc).Speed = Val(General_Var_Get(StreamFile, Val(loopc), "Speed"))
            StreamData(loopc).NumGrhs = General_Var_Get(StreamFile, Val(loopc), "NumGrhs")
           
            ReDim StreamData(loopc).grh_list(1 To StreamData(loopc).NumGrhs)
            GrhListing = General_Var_Get(StreamFile, Val(loopc), "Grh_List")
           
            For i = 1 To StreamData(loopc).NumGrhs
                StreamData(loopc).grh_list(i) = General_Field_Read(str(i), GrhListing, 44)
            Next i
            StreamData(loopc).grh_list(i - 1) = StreamData(loopc).grh_list(i - 1)
            For ColorSet = 1 To 4
                TempSet = General_Var_Get(StreamFile, Val(loopc), "ColorSet" & ColorSet)
                StreamData(loopc).colortint(ColorSet - 1).r = General_Field_Read(1, TempSet, 44)
                StreamData(loopc).colortint(ColorSet - 1).g = General_Field_Read(2, TempSet, 44)
                StreamData(loopc).colortint(ColorSet - 1).B = General_Field_Read(3, TempSet, 44)
            Next ColorSet
        Next loopc
 
End Sub
'*****************************************************************
'************************Generar particulas en Cuerpos************
Public Function General_Char_Particle_Create(ByVal ParticulaInd As Long, ByVal char_index As Integer, Optional ByVal particle_life As Long = 0) As Long
On Error Resume Next

If ParticulaInd <= 0 Then Exit Function

Dim rgb_list(0 To 3) As Long
rgb_list(0) = RGB(StreamData(ParticulaInd).colortint(0).r, StreamData(ParticulaInd).colortint(0).g, StreamData(ParticulaInd).colortint(0).B)
rgb_list(1) = RGB(StreamData(ParticulaInd).colortint(1).r, StreamData(ParticulaInd).colortint(1).g, StreamData(ParticulaInd).colortint(1).B)
rgb_list(2) = RGB(StreamData(ParticulaInd).colortint(2).r, StreamData(ParticulaInd).colortint(2).g, StreamData(ParticulaInd).colortint(2).B)
rgb_list(3) = RGB(StreamData(ParticulaInd).colortint(3).r, StreamData(ParticulaInd).colortint(3).g, StreamData(ParticulaInd).colortint(3).B)

General_Char_Particle_Create = Char_Particle_Group_Create(char_index, StreamData(ParticulaInd).grh_list, rgb_list(), StreamData(ParticulaInd).NumOfParticles, ParticulaInd, _
    StreamData(ParticulaInd).AlphaBlend, IIf(particle_life = 0, StreamData(ParticulaInd).life_counter, particle_life), StreamData(ParticulaInd).Speed, , StreamData(ParticulaInd).x1, StreamData(ParticulaInd).y1, StreamData(ParticulaInd).angle, _
    StreamData(ParticulaInd).vecx1, StreamData(ParticulaInd).vecx2, StreamData(ParticulaInd).vecy1, StreamData(ParticulaInd).vecy2, _
    StreamData(ParticulaInd).life1, StreamData(ParticulaInd).life2, StreamData(ParticulaInd).friction, StreamData(ParticulaInd).spin_speedL, _
    StreamData(ParticulaInd).gravity, StreamData(ParticulaInd).grav_strength, StreamData(ParticulaInd).bounce_strength, StreamData(ParticulaInd).x2, _
    StreamData(ParticulaInd).y2, StreamData(ParticulaInd).XMove, StreamData(ParticulaInd).move_x1, StreamData(ParticulaInd).move_x2, StreamData(ParticulaInd).move_y1, _
    StreamData(ParticulaInd).move_y2, StreamData(ParticulaInd).YMove, StreamData(ParticulaInd).spin_speedH, StreamData(ParticulaInd).spin, _
    StreamData(ParticulaInd).Radio)

End Function
'*******************************************************************
'******************Generar particulas en el mapa********************
Public Function General_Particle_Create(ByVal ParticulaInd As Long, ByVal X As Integer, ByVal Y As Integer, Optional ByVal particle_life As Long = 0) As Long
   
Dim rgb_list(0 To 3) As Long
rgb_list(0) = RGB(StreamData(ParticulaInd).colortint(0).r, StreamData(ParticulaInd).colortint(0).g, StreamData(ParticulaInd).colortint(0).B)
rgb_list(1) = RGB(StreamData(ParticulaInd).colortint(1).r, StreamData(ParticulaInd).colortint(1).g, StreamData(ParticulaInd).colortint(1).B)
rgb_list(2) = RGB(StreamData(ParticulaInd).colortint(2).r, StreamData(ParticulaInd).colortint(2).g, StreamData(ParticulaInd).colortint(2).B)
rgb_list(3) = RGB(StreamData(ParticulaInd).colortint(3).r, StreamData(ParticulaInd).colortint(3).g, StreamData(ParticulaInd).colortint(3).B)
 
General_Particle_Create = Particle_Group_Create(X, Y, StreamData(ParticulaInd).grh_list, rgb_list(), StreamData(ParticulaInd).NumOfParticles, ParticulaInd, _
    StreamData(ParticulaInd).AlphaBlend, IIf(particle_life = 0, StreamData(ParticulaInd).life_counter, particle_life), StreamData(ParticulaInd).Speed, , StreamData(ParticulaInd).x1, StreamData(ParticulaInd).y1, StreamData(ParticulaInd).angle, _
    StreamData(ParticulaInd).vecx1, StreamData(ParticulaInd).vecx2, StreamData(ParticulaInd).vecy1, StreamData(ParticulaInd).vecy2, _
    StreamData(ParticulaInd).life1, StreamData(ParticulaInd).life2, StreamData(ParticulaInd).friction, StreamData(ParticulaInd).spin_speedL, _
    StreamData(ParticulaInd).gravity, StreamData(ParticulaInd).grav_strength, StreamData(ParticulaInd).bounce_strength, StreamData(ParticulaInd).x2, _
    StreamData(ParticulaInd).y2, StreamData(ParticulaInd).XMove, StreamData(ParticulaInd).move_x1, StreamData(ParticulaInd).move_x2, StreamData(ParticulaInd).move_y1, _
    StreamData(ParticulaInd).move_y2, StreamData(ParticulaInd).YMove, StreamData(ParticulaInd).spin_speedH, StreamData(ParticulaInd).spin, , , , _
    StreamData(ParticulaInd).Radio)
End Function
'*******************************************************************
