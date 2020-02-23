Attribute VB_Name = "Trabajo"
Option Explicit

Private Const ENERGIA_TRABAJO_HERRERO As Byte = 2
Private Const ENERGIA_TRABAJO_NOHERRERO As Byte = 6


Public Sub DoPermanecerOculto(ByVal Userindex As Integer)
'********************************************************
'Autor: Nacho (Integer)
'Last Modif: 28/01/2007
'Chequea si ya debe mostrarse
'Pablo (ToxicWaste): Cambie los ordenes de prioridades porque sino no andaba.
'********************************************************

UserList(Userindex).Counters.TiempoOculto = UserList(Userindex).Counters.TiempoOculto - 1
If UserList(Userindex).Counters.TiempoOculto <= 0 Then
    
    UserList(Userindex).Counters.TiempoOculto = IntervaloOculto
    If UserList(Userindex).clase = eClass.Hunter And UserList(Userindex).Stats.UserSkills(eSkill.Ocultarse) > 90 Then
        If UserList(Userindex).Invent.ArmourEqpObjIndex = 648 Or UserList(Userindex).Invent.ArmourEqpObjIndex = 360 Then
            Exit Sub
        End If
    End If
    UserList(Userindex).Counters.TiempoOculto = 0
    UserList(Userindex).flags.Oculto = 0
    Call WriteTimeInvi(Userindex, 0)
    If UserList(Userindex).flags.invisible = 0 Then
        Call WriteConsoleMsg(Userindex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
        Call SetInvisible(Userindex, UserList(Userindex).Char.CharIndex, False)
        'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, False))
    End If
End If

Exit Sub

Errhandler:
    Call LogError("Error en Sub DoPermanecerOculto")


End Sub

Public Sub DoOcultarse(ByVal Userindex As Integer)
'Pablo (ToxicWaste): No olvidar agregar IntervaloOculto=500 al Server.ini.
'Modifique la fórmula y ahora anda bien.
On Error GoTo Errhandler

Dim Suerte As Double
Dim res As Integer
Dim Skill As Integer

Skill = UserList(Userindex).Stats.UserSkills(eSkill.Ocultarse)

Suerte = (((0.000002 * Skill - 0.0002) * Skill + 0.0064) * Skill + 0.1124) * 100

res = RandomNumber(1, 100)

If res <= Suerte Then

    UserList(Userindex).flags.Oculto = 1
    Suerte = (-0.000001 * (100 - Skill) ^ 3)
    Suerte = Suerte + (0.00009229 * (100 - Skill) ^ 2)
    Suerte = Suerte + (-0.0088 * (100 - Skill))
    Suerte = Suerte + (0.9571)
    Suerte = Suerte * IntervaloOculto
    UserList(Userindex).Counters.TiempoOculto = Suerte
  
    Call SetInvisible(Userindex, UserList(Userindex).Char.CharIndex, True)
    'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, True))

    Call WriteConsoleMsg(Userindex, "¡Te has escondido entre las sombras!", FontTypeNames.FONTTYPE_INFO)
    Call SubirSkill(Userindex, Ocultarse)
Else
    '[CDT 17-02-2004]
    If Not UserList(Userindex).flags.UltimoMensaje = 4 Then
        Call WriteConsoleMsg(Userindex, "¡No has logrado esconderte!", FontTypeNames.FONTTYPE_INFO)
        UserList(Userindex).flags.UltimoMensaje = 4
    End If
    '[/CDT]
End If

UserList(Userindex).Counters.Ocultando = UserList(Userindex).Counters.Ocultando + 1

Exit Sub

Errhandler:
    Call LogError("Error en Sub DoOcultarse")

End Sub

Public Sub DoNavega(ByVal Userindex As Integer, ByRef Barco As ObjData, ByVal Slot As Integer)

Dim ModNave As Long

With UserList(Userindex)

ModNave = ModNavegacion(.clase)

If .Stats.UserSkills(eSkill.Navegacion) / ModNave < Barco.MinSkill Then
    Call WriteConsoleMsg(Userindex, "No tenes suficientes conocimientos para usar este barco.", FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(Userindex, "Para usar este barco necesitas " & Barco.MinSkill * ModNave & " puntos en navegacion.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If UserList(Userindex).flags.Metamorfosis = 1 Then 'Metamorfosis
    Call WriteConsoleMsg(Userindex, "No puedes navegar mientras estas metamorfoseado.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

.Invent.BarcoObjIndex = .Invent.Object(Slot).ObjIndex
.Invent.BarcoSlot = Slot

If .flags.Navegando = 0 Then
    
    .Char.Head = 0
    
    If .flags.Muerto = 0 Then
        '(Nacho)
        If .Faccion.ArmadaReal = 1 Then
            .Char.Body = iFragataReal
        ElseIf .Faccion.FuerzasCaos = 1 Then
            .Char.Body = iFragataCaos
        Else
            If criminal(Userindex) Then
                If Barco.Ropaje = iBarca Then .Char.Body = iBarcaPk
                If Barco.Ropaje = iGalera Then .Char.Body = iGaleraPk
                If Barco.Ropaje = iGaleon Then .Char.Body = iGaleonPk
            Else
                If Barco.Ropaje = iBarca Then .Char.Body = iBarcaCiuda
                If Barco.Ropaje = iGalera Then .Char.Body = iGaleraCiuda
                If Barco.Ropaje = iGaleon Then .Char.Body = iGaleonCiuda
            End If
        End If
    Else
        .Char.Body = iFragataFantasmal
    End If
    
    .Char.ShieldAnim = NingunEscudo
    .Char.WeaponAnim = NingunArma
    .Char.CascoAnim = NingunCasco
    .flags.Navegando = 1
    
Else
    
        If Not ((LegalPos(.Pos.map, .Pos.X - 1, .Pos.Y, False, True) _
            Or LegalPos(.Pos.map, .Pos.X, .Pos.Y - 1, False, True) _
            Or LegalPos(.Pos.map, .Pos.X + 1, .Pos.Y, False, True) _
            Or LegalPos(.Pos.map, .Pos.X, .Pos.Y + 1, False, True)) _
            And .flags.Navegando = 1) Or _
            .flags.Navegando = 0 Then
                Call WriteConsoleMsg(Userindex, "¡Debes aproximarte a la costa para dejar el barco!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
        End If
    
    .flags.Navegando = 0
    
    If .flags.Muerto = 0 Then
        .Char.Head = .OrigChar.Head
        
        If .Invent.ArmourEqpObjIndex > 0 Then
            .Char.Body = ObjData(.Invent.ArmourEqpObjIndex).Ropaje
        Else
            Call DarCuerpoDesnudo(Userindex)
        End If
        
        If .Invent.EscudoEqpObjIndex > 0 Then _
            .Char.ShieldAnim = ObjData(.Invent.EscudoEqpObjIndex).ShieldAnim
        If .Invent.WeaponEqpObjIndex > 0 Then _
            .Char.WeaponAnim = ObjData(.Invent.WeaponEqpObjIndex).WeaponAnim
        If .Invent.CascoEqpObjIndex > 0 Then _
            .Char.CascoAnim = ObjData(.Invent.CascoEqpObjIndex).CascoAnim
    Else
        .Char.Body = iCuerpoMuerto
        .Char.Head = iCabezaMuerto
        .Char.ShieldAnim = NingunEscudo
        .Char.WeaponAnim = NingunArma
        .Char.CascoAnim = NingunCasco
    End If
End If

Call ChangeUserChar(Userindex, .Char.Body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.Aura)
Call WriteNavigateToggle(Userindex)
End With
End Sub
Public Sub DoEquita(ByVal Userindex As Integer, ByRef Montura As ObjData, ByVal Slot As Integer)

Dim ModEqui As Long

ModEqui = ModEquitacion(UserList(Userindex).clase)

If UserList(Userindex).Stats.UserSkills(Equitacion) / ModEqui < Montura.MinSkill Then
    Call WriteConsoleMsg(Userindex, "Para usar esta montura necesitas " & Montura.MinSkill * ModEqui & " puntos en equitación.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If UserList(Userindex).flags.Navegando = 1 Then
    Call WriteConsoleMsg(Userindex, "No puedes utilizar la montura mientras navegas !!", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If Not UserList(Userindex).flags.Equitando = 1 Then
    If MapData(UserList(Userindex).Pos.map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).trigger = BAJOTECHO Then
        Call WriteConsoleMsg(Userindex, "No puedes utilizar la montura ahi!", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
End If

If UserList(Userindex).flags.Metamorfosis = 1 Then 'Metamorfosis
    Call WriteConsoleMsg(Userindex, "No puedes montar mientras estas metamorfoseado.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

UserList(Userindex).Invent.MonturaObjIndex = UserList(Userindex).Invent.Object(Slot).ObjIndex
UserList(Userindex).Invent.MonturaSlot = Slot

If UserList(Userindex).flags.Equitando = 0 Then
    UserList(Userindex).Char.Head = 0
    If UserList(Userindex).flags.Muerto = 0 Then
        UserList(Userindex).Char.Body = Montura.Ropaje
    Else
        UserList(Userindex).Char.Body = iCuerpoMuerto
        UserList(Userindex).Char.Head = iCabezaMuerto
    End If
    UserList(Userindex).Char.Head = UserList(Userindex).OrigChar.Head
    UserList(Userindex).Char.WeaponAnim = NingunArma
    UserList(Userindex).Char.CascoAnim = UserList(Userindex).Char.CascoAnim
    UserList(Userindex).flags.Equitando = 1
Else
    UserList(Userindex).flags.Equitando = 0
    If UserList(Userindex).flags.Muerto = 0 Then
        UserList(Userindex).Char.Head = UserList(Userindex).OrigChar.Head
        If UserList(Userindex).Invent.ArmourEqpObjIndex > 0 Then
            UserList(Userindex).Char.Body = ObjData(UserList(Userindex).Invent.ArmourEqpObjIndex).Ropaje
        Else
            Call DarCuerpoDesnudo(Userindex)
        End If
        If UserList(Userindex).Invent.EscudoEqpObjIndex > 0 Then UserList(Userindex).Char.ShieldAnim = ObjData(UserList(Userindex).Invent.EscudoEqpObjIndex).ShieldAnim
        If UserList(Userindex).Invent.WeaponEqpObjIndex > 0 Then UserList(Userindex).Char.WeaponAnim = ObjData(UserList(Userindex).Invent.WeaponEqpObjIndex).WeaponAnim
        If UserList(Userindex).Invent.CascoEqpObjIndex > 0 Then UserList(Userindex).Char.CascoAnim = ObjData(UserList(Userindex).Invent.CascoEqpObjIndex).CascoAnim
    Else
        UserList(Userindex).Char.Body = iCuerpoMuerto
        UserList(Userindex).Char.Head = iCabezaMuerto
        UserList(Userindex).Char.ShieldAnim = NingunEscudo
        UserList(Userindex).Char.WeaponAnim = NingunArma
        UserList(Userindex).Char.CascoAnim = NingunCasco
    End If
End If

Call ChangeUserChar(Userindex, UserList(Userindex).Char.Body, UserList(Userindex).Char.Head, UserList(Userindex).Char.heading, UserList(Userindex).Char.WeaponAnim, UserList(Userindex).Char.ShieldAnim, UserList(Userindex).Char.CascoAnim, UserList(Userindex).Char.Aura)
Call WriteEquitandoToggle(Userindex)

End Sub
Function ModEquitacion(ByVal clase As String) As Integer

Select Case UCase$(clase)
    Case "CLERIGO"
        ModEquitacion = 1
    Case Else
        ModEquitacion = 1.5
End Select

End Function

Public Sub FundirMineral(ByVal Userindex As Integer)

On Error GoTo Errhandler

If UserList(Userindex).flags.TargetObjInvIndex > 0 Then
   
   If ObjData(UserList(Userindex).flags.TargetObjInvIndex).OBJType = eOBJType.otMinerales And ObjData(UserList(Userindex).flags.TargetObjInvIndex).MinSkill <= UserList(Userindex).Stats.UserSkills(eSkill.Mineria) / ModFundicion(UserList(Userindex).clase) Then
        Call DoLingotes(Userindex)
   Else
        Call WriteConsoleMsg(Userindex, "No tenes conocimientos de mineria suficientes para trabajar este mineral.", FontTypeNames.FONTTYPE_INFO)
   End If

End If

Exit Sub

Errhandler:
    Call LogError("Error en FundirMineral. Error " & Err.Number & " : " & Err.description)

End Sub
Function TieneObjetos(ByVal ItemIndex As Integer, ByVal cant As Integer, ByVal Userindex As Integer) As Boolean
'Call LogTarea("Sub TieneObjetos")

Dim i As Integer
Dim Total As Long
For i = 1 To MAX_INVENTORY_SLOTS
    If UserList(Userindex).Invent.Object(i).ObjIndex = ItemIndex Then
        Total = Total + UserList(Userindex).Invent.Object(i).amount
    End If
Next i

If cant <= Total Then
    TieneObjetos = True
    Exit Function
End If
        
End Function

Public Sub QuitarObjetos(ByVal ItemIndex As Integer, ByVal cant As Integer, ByVal Userindex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 05/08/09
'05/08/09: Pato - Cambie la funcion a procedimiento ya que se usa como procedimiento siempre, y fixie el bug 2788199
'***************************************************

'Call LogTarea("Sub QuitarObjetos")

Dim i As Integer
For i = 1 To MAX_INVENTORY_SLOTS
    With UserList(Userindex).Invent.Object(i)
        If .ObjIndex = ItemIndex Then
            If .amount <= cant And .Equipped = 1 Then Call Desequipar(Userindex, i)
            
            .amount = .amount - cant
            If .amount <= 0 Then
                cant = Abs(.amount)
                UserList(Userindex).Invent.NroItems = UserList(Userindex).Invent.NroItems - 1
                .amount = 0
                .ObjIndex = 0
            Else
                cant = 0
            End If
            
            Call UpdateUserInv(False, Userindex, i)
            
            If cant = 0 Then Exit Sub
        End If
    End With
Next i

End Sub

Sub HerreroQuitarMateriales(ByVal Userindex As Integer, ByVal ItemIndex As Integer, ByVal Cantidad As Integer)
    If ObjData(ItemIndex).LingH > 0 Then Call QuitarObjetos(LingoteHierro, ObjData(ItemIndex).LingH * Cantidad, Userindex)
    If ObjData(ItemIndex).LingP > 0 Then Call QuitarObjetos(LingotePlata, ObjData(ItemIndex).LingP * Cantidad, Userindex)
    If ObjData(ItemIndex).LingO > 0 Then Call QuitarObjetos(LingoteOro, ObjData(ItemIndex).LingO * Cantidad, Userindex)
End Sub

Sub CarpinteroQuitarMateriales(ByVal Userindex As Integer, ByVal ItemIndex As Integer, ByVal Cantidad As Integer)
    If ObjData(ItemIndex).Madera > 0 Then Call QuitarObjetos(Leña, ObjData(ItemIndex).Madera * Cantidad, Userindex)
End Sub

Function CarpinteroTieneMateriales(ByVal Userindex As Integer, ByVal ItemIndex As Integer, ByVal Cantidad As Integer) As Boolean
    
    If ObjData(ItemIndex).Madera > 0 Then
            If Not TieneObjetos(Leña, ObjData(ItemIndex).Madera * Cantidad, Userindex) Then
                    Call WriteConsoleMsg(Userindex, "No tenes suficientes madera.", FontTypeNames.FONTTYPE_INFO)
                    CarpinteroTieneMateriales = False
                    Exit Function
            End If
    End If
    
    CarpinteroTieneMateriales = True

End Function
 
Function HerreroTieneMateriales(ByVal Userindex As Integer, ByVal ItemIndex As Integer, ByVal Cantidad As Integer) As Boolean
    If ObjData(ItemIndex).LingH > 0 Then
            If Not TieneObjetos(LingoteHierro, ObjData(ItemIndex).LingH * Cantidad, Userindex) Then
                    Call WriteConsoleMsg(Userindex, "No tenes suficientes lingotes de hierro.", FontTypeNames.FONTTYPE_INFO)
                    HerreroTieneMateriales = False
                    Exit Function
            End If
    End If
    If ObjData(ItemIndex).LingP > 0 Then
            If Not TieneObjetos(LingotePlata * Cantidad, ObjData(ItemIndex).LingP, Userindex) Then
                    Call WriteConsoleMsg(Userindex, "No tenes suficientes lingotes de plata.", FontTypeNames.FONTTYPE_INFO)
                    HerreroTieneMateriales = False
                    Exit Function
            End If
    End If
    If ObjData(ItemIndex).LingO > 0 Then
            If Not TieneObjetos(LingoteOro * Cantidad, ObjData(ItemIndex).LingO, Userindex) Then
                    Call WriteConsoleMsg(Userindex, "No tenes suficientes lingotes de oro.", FontTypeNames.FONTTYPE_INFO)
                    HerreroTieneMateriales = False
                    Exit Function
            End If
    End If
    HerreroTieneMateriales = True
End Function

Public Function PuedeConstruir(ByVal Userindex As Integer, ByVal ItemIndex As Integer, ByVal Cantidad As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 24/08/2009
'24/08/2008: ZaMa - Validates if the player has the required skill
'***************************************************
PuedeConstruir = HerreroTieneMateriales(Userindex, ItemIndex, Cantidad) And _
                    Round(UserList(Userindex).Stats.UserSkills(eSkill.Herreria) / ModHerreriA(UserList(Userindex).clase), 0) >= ObjData(ItemIndex).SkHerreria
End Function

Public Function PuedeConstruirHerreria(ByVal ItemIndex As Integer) As Boolean
Dim i As Long

For i = 1 To UBound(ArmasHerrero)
    If ArmasHerrero(i) = ItemIndex Then
        PuedeConstruirHerreria = True
        Exit Function
    End If
Next i
For i = 1 To UBound(ArmadurasHerrero)
    If ArmadurasHerrero(i) = ItemIndex Then
        PuedeConstruirHerreria = True
        Exit Function
    End If
Next i
PuedeConstruirHerreria = False
End Function


Public Sub HerreroConstruirItem(ByVal Userindex As Integer, ByVal ItemIndex As Integer, ByVal Cantidad As Integer)

If PuedeConstruir(Userindex, ItemIndex, Cantidad) And PuedeConstruirHerreria(ItemIndex) Then
    
    'Sacamos energía
    If UserList(Userindex).clase = eClass.trabajador Then
        'Chequeamos que tenga los puntos antes de sacarselos
        If UserList(Userindex).Stats.MinSta >= ENERGIA_TRABAJO_HERRERO Then
            UserList(Userindex).Stats.MinSta = UserList(Userindex).Stats.MinSta - ENERGIA_TRABAJO_HERRERO
            Call WriteUpdateSta(Userindex)
        Else
            Call WriteConsoleMsg(Userindex, "No tienes suficiente energía.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    Else
        'Chequeamos que tenga los puntos antes de sacarselos
        If UserList(Userindex).Stats.MinSta >= ENERGIA_TRABAJO_NOHERRERO Then
            UserList(Userindex).Stats.MinSta = UserList(Userindex).Stats.MinSta - ENERGIA_TRABAJO_NOHERRERO
            Call WriteUpdateSta(Userindex)
        Else
            Call WriteConsoleMsg(Userindex, "No tienes suficiente energía.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    End If
    
    Call HerreroQuitarMateriales(Userindex, ItemIndex, Cantidad)
    ' AGREGAR FX
    If ObjData(ItemIndex).OBJType = eOBJType.otWeapon Then
        Call WriteConsoleMsg(Userindex, "Has construido el arma!.", FontTypeNames.FONTTYPE_INFO)
    ElseIf ObjData(ItemIndex).OBJType = eOBJType.otESCUDO Then
        Call WriteConsoleMsg(Userindex, "Has construido el escudo!.", FontTypeNames.FONTTYPE_INFO)
    ElseIf ObjData(ItemIndex).OBJType = eOBJType.otCASCO Then
        Call WriteConsoleMsg(Userindex, "Has construido el casco!.", FontTypeNames.FONTTYPE_INFO)
    ElseIf ObjData(ItemIndex).OBJType = eOBJType.otArmadura Then
        Call WriteConsoleMsg(Userindex, "Has construido la armadura!.", FontTypeNames.FONTTYPE_INFO)
    End If
    Dim MiObj As Obj
    MiObj.amount = Cantidad
    MiObj.ObjIndex = ItemIndex
    If Not MeterItemEnInventario(Userindex, MiObj) Then
        Call TirarItemAlPiso(UserList(Userindex).Pos, MiObj)
    End If
    
    'Log de construcción de Items. Pablo (ToxicWaste) 10/09/07
    If ObjData(MiObj.ObjIndex).Log = 1 Then
        Call LogDesarrollo(UserList(Userindex).Name & " ha construído " & MiObj.amount & " " & ObjData(MiObj.ObjIndex).Name)
    End If
    
    Call SubirSkill(Userindex, Herreria)
    Call UpdateUserInv(True, Userindex, 0)
    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(MARTILLOHERRERO, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

    UserList(Userindex).Reputacion.PlebeRep = UserList(Userindex).Reputacion.PlebeRep + vlProleta
    If UserList(Userindex).Reputacion.PlebeRep > MAXREP Then _
        UserList(Userindex).Reputacion.PlebeRep = MAXREP

    UserList(Userindex).Counters.Trabajando = UserList(Userindex).Counters.Trabajando + 1
End If
End Sub

Public Function PuedeConstruirCarpintero(ByVal ItemIndex As Integer) As Boolean
Dim i As Long

For i = 1 To UBound(ObjCarpintero)
    If ObjCarpintero(i) = ItemIndex Then
        PuedeConstruirCarpintero = True
        Exit Function
    End If
Next i
PuedeConstruirCarpintero = False

End Function

Public Sub CarpinteroConstruirItem(ByVal Userindex As Integer, ByVal ItemIndex As Integer, ByVal Cantidad As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 24/08/2009
'24/08/2008: ZaMa - Validates if the player has the required skill
'***************************************************
If CarpinteroTieneMateriales(Userindex, ItemIndex, Cantidad) And _
   Round(UserList(Userindex).Stats.UserSkills(eSkill.Carpinteria) \ ModCarpinteria(UserList(Userindex).clase), 0) >= _
   ObjData(ItemIndex).SkCarpinteria And _
   PuedeConstruirCarpintero(ItemIndex) And _
   UserList(Userindex).Invent.WeaponEqpObjIndex = SERRUCHO_CARPINTERO Then
   
    'Sacamos energía
    If UserList(Userindex).clase = eClass.trabajador Then
        'Chequeamos que tenga los puntos antes de sacarselos
        If UserList(Userindex).Stats.MinSta >= ENERGIA_TRABAJO_HERRERO Then
            UserList(Userindex).Stats.MinSta = UserList(Userindex).Stats.MinSta - ENERGIA_TRABAJO_HERRERO
            Call WriteUpdateSta(Userindex)
        Else
            Call WriteConsoleMsg(Userindex, "No tienes suficiente energía.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    Else
        'Chequeamos que tenga los puntos antes de sacarselos
        If UserList(Userindex).Stats.MinSta >= ENERGIA_TRABAJO_NOHERRERO Then
            UserList(Userindex).Stats.MinSta = UserList(Userindex).Stats.MinSta - ENERGIA_TRABAJO_NOHERRERO
            Call WriteUpdateSta(Userindex)
        Else
            Call WriteConsoleMsg(Userindex, "No tienes suficiente energía.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    End If
    
    Call CarpinteroQuitarMateriales(Userindex, ItemIndex, Cantidad)
    Call WriteConsoleMsg(Userindex, "Has construido el objeto!.", FontTypeNames.FONTTYPE_INFO)
    
    Dim MiObj As Obj
    MiObj.amount = Cantidad
    MiObj.ObjIndex = ItemIndex
    If Not MeterItemEnInventario(Userindex, MiObj) Then
                    Call TirarItemAlPiso(UserList(Userindex).Pos, MiObj)
    End If
    
    'Log de construcción de Items. Pablo (ToxicWaste) 10/09/07
    If ObjData(MiObj.ObjIndex).Log = 1 Then
        Call LogDesarrollo(UserList(Userindex).Name & " ha construído " & MiObj.amount & " " & ObjData(MiObj.ObjIndex).Name)
    End If
    
    Call SubirSkill(Userindex, Carpinteria)
    Call UpdateUserInv(True, Userindex, 0)
    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(LABUROCARPINTERO, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))


    UserList(Userindex).Reputacion.PlebeRep = UserList(Userindex).Reputacion.PlebeRep + vlProleta
    If UserList(Userindex).Reputacion.PlebeRep > MAXREP Then _
        UserList(Userindex).Reputacion.PlebeRep = MAXREP

    UserList(Userindex).Counters.Trabajando = UserList(Userindex).Counters.Trabajando + 1

End If
End Sub

Private Function MineralesParaLingote(ByVal Lingote As iMinerales) As Integer
    Select Case Lingote
        Case iMinerales.HierroCrudo
            MineralesParaLingote = 14
        Case iMinerales.PlataCruda
            MineralesParaLingote = 20
        Case iMinerales.OroCrudo
            MineralesParaLingote = 35
        Case Else
            MineralesParaLingote = 10000
    End Select
End Function


Public Sub DoLingotes(ByVal Userindex As Integer)
'    Call LogTarea("Sub DoLingotes")
    Dim Slot As Integer
    Dim obji As Integer

    Slot = UserList(Userindex).flags.TargetObjInvSlot
    obji = UserList(Userindex).Invent.Object(Slot).ObjIndex
    
    If UserList(Userindex).Invent.Object(Slot).amount < MineralesParaLingote(obji) Or _
        ObjData(obji).OBJType <> eOBJType.otMinerales Then
            Call WriteConsoleMsg(Userindex, "No tienes mas minerales para fundir... Dejas de trabajar.", FontTypeNames.FONTTYPE_INFO)
            UserList(Userindex).flags.Makro = 0
            Call WriteConsoleMsg(Userindex, "Macro de Fundición asistido desactivado.", FontTypeNames.FONTTYPE_VENENO)
            Exit Sub
    End If
    
    UserList(Userindex).Invent.Object(Slot).amount = UserList(Userindex).Invent.Object(Slot).amount - MineralesParaLingote(obji)
    If UserList(Userindex).Invent.Object(Slot).amount < 1 Then
        UserList(Userindex).Invent.Object(Slot).amount = 0
        UserList(Userindex).Invent.Object(Slot).ObjIndex = 0
    End If
    
    Dim MiObj As Obj
    MiObj.amount = 1
    MiObj.ObjIndex = ObjData(UserList(Userindex).flags.TargetObjInvIndex).LingoteIndex
    If Not MeterItemEnInventario(Userindex, MiObj) Then
        Call TirarItemAlPiso(UserList(Userindex).Pos, MiObj)
    End If
    Call UpdateUserInv(False, Userindex, Slot)
    Call WriteConsoleMsg(Userindex, "¡Has obtenido un lingote!", FontTypeNames.FONTTYPE_INFO)

    UserList(Userindex).Counters.Trabajando = UserList(Userindex).Counters.Trabajando + 1
End Sub

Function ModNavegacion(ByVal clase As eClass) As Single

Select Case clase
    Case eClass.Pirat
        ModNavegacion = 1
    Case eClass.trabajador
        ModNavegacion = 1.2
    Case Else
        ModNavegacion = 2.3
End Select

End Function


Function ModFundicion(ByVal clase As eClass) As Single

Select Case clase
    Case eClass.trabajador
        ModFundicion = 1
    Case Else
        ModFundicion = 3
End Select

End Function

Function ModCarpinteria(ByVal clase As eClass) As Integer

Select Case clase
    Case eClass.trabajador
        ModCarpinteria = 1
    Case Else
        ModCarpinteria = 3
End Select

End Function

Function ModHerreriA(ByVal clase As eClass) As Single
Select Case clase
    Case eClass.trabajador
        ModHerreriA = 1
    Case Else
        ModHerreriA = 4
End Select

End Function

Function ModDomar(ByVal clase As eClass) As Integer
    Select Case clase
        Case eClass.Druid
            ModDomar = 6
        Case eClass.Hunter
            ModDomar = 6
        Case eClass.Cleric
            ModDomar = 7
        Case Else
            ModDomar = 10
    End Select
End Function

Function FreeMascotaIndex(ByVal Userindex As Integer) As Integer
'***************************************************
'Author: Unknown
'Last Modification: 02/03/09
'02/03/09: ZaMa - Busca un indice libre de mascotas, revisando los types y no los indices de los npcs
'***************************************************
    Dim j As Integer
    For j = 1 To MAXMASCOTAS
        If UserList(Userindex).MascotasType(j) = 0 Then
            FreeMascotaIndex = j
            Exit Function
        End If
    Next j
End Function

Sub DoDomar(ByVal Userindex As Integer, ByVal NpcIndex As Integer)
'***************************************************
'Author: Nacho (Integer)
'Last Modification: 02/03/2009
'12/15/2008: ZaMa - Limits the number of the same type of pet to 2.
'02/03/2009: ZaMa - Las criaturas domadas en zona segura, esperan afuera (desaparecen).
'***************************************************

On Error GoTo Errhandler

Dim puntosDomar As Integer
Dim puntosRequeridos As Integer
Dim CanStay As Boolean
Dim petType As Integer
Dim NroPets As Integer


If Npclist(NpcIndex).MaestroUser = Userindex Then
    Call WriteConsoleMsg(Userindex, "Ya domaste a esa criatura.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If UserList(Userindex).NroMascotas < MAXMASCOTAS Then
    
    If Npclist(NpcIndex).MaestroNpc > 0 Or Npclist(NpcIndex).MaestroUser > 0 Then
        Call WriteConsoleMsg(Userindex, "La criatura ya tiene amo.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If Not PuedeDomarMascota(Userindex, NpcIndex) Then
        Call WriteConsoleMsg(Userindex, "No puedes domar mas de dos criaturas del mismo tipo.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    puntosDomar = CInt(UserList(Userindex).Stats.UserAtributos(eAtributos.Carisma)) * CInt(UserList(Userindex).Stats.UserSkills(eSkill.Domar))
    If UserList(Userindex).Invent.AnilloEqpObjIndex = FLAUTAMAGICA Then
        puntosRequeridos = Npclist(NpcIndex).flags.Domable * 0.8
    Else
        puntosRequeridos = Npclist(NpcIndex).flags.Domable
    End If
    
    If puntosRequeridos <= puntosDomar And RandomNumber(1, 5) = 1 Then
        Dim index As Integer
        UserList(Userindex).NroMascotas = UserList(Userindex).NroMascotas + 1
        index = FreeMascotaIndex(Userindex)
        UserList(Userindex).MascotasIndex(index) = NpcIndex
        UserList(Userindex).MascotasType(index) = Npclist(NpcIndex).Numero
        
        Npclist(NpcIndex).MaestroUser = Userindex
        
        Call FollowAmo(NpcIndex)
        Call ReSpawnNpc(Npclist(NpcIndex))
        
        Call WriteConsoleMsg(Userindex, "La criatura te ha aceptado como su amo.", FontTypeNames.FONTTYPE_INFO)
        
        ' Es zona segura?
        CanStay = (MapInfo(UserList(Userindex).Pos.map).Pk = True)
        
        If Not CanStay Then
            petType = Npclist(NpcIndex).Numero
            NroPets = UserList(Userindex).NroMascotas
            
            Call QuitarNPC(NpcIndex)
            
            UserList(Userindex).MascotasType(index) = petType
            UserList(Userindex).NroMascotas = NroPets
            
            Call WriteConsoleMsg(Userindex, "No se permiten mascotas en zona segura. Éstas te esperarán afuera.", FontTypeNames.FONTTYPE_INFO)
        End If

    Else
        If Not UserList(Userindex).flags.UltimoMensaje = 5 Then
            Call WriteConsoleMsg(Userindex, "No has logrado domar la criatura.", FontTypeNames.FONTTYPE_INFO)
            UserList(Userindex).flags.UltimoMensaje = 5
        End If
    End If
    
    'Entreno domar. Es un 30% más dificil si no sos druida.
    If UserList(Userindex).clase = eClass.Druid Or (RandomNumber(1, 3) < 3) Then
        Call SubirSkill(Userindex, Domar)
    End If
Else
    Call WriteConsoleMsg(Userindex, "No puedes controlar más criaturas.", FontTypeNames.FONTTYPE_INFO)
End If

Exit Sub

Errhandler:
    Call LogError("Error en DoDomar. Error " & Err.Number & " : " & Err.description)

End Sub

''
' Checks if the user can tames a pet.
'
' @param integer userIndex The user id from who wants tame the pet.
' @param integer NPCindex The index of the npc to tome.
' @return boolean True if can, false if not.
Private Function PuedeDomarMascota(ByVal Userindex As Integer, ByVal NpcIndex As Integer) As Boolean
'***************************************************
'Author: ZaMa
'This function checks how many NPCs of the same type have
'been tamed by the user.
'Returns True if that amount is less than two.
'***************************************************
    Dim i As Long
    Dim numMascotas As Long
    
    For i = 1 To MAXMASCOTAS
        If UserList(Userindex).MascotasType(i) = Npclist(NpcIndex).Numero Then
            numMascotas = numMascotas + 1
        End If
    Next i
    
    If numMascotas <= 1 Then PuedeDomarMascota = True
    
End Function

Sub DoAdminInvisible(ByVal Userindex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 13/07/2009
'Makes an admin invisible o visible.
'13/07/2009: ZaMa - Now invisible admins' chars are erased from all clients, except from themselves.
'***************************************************
    
    With UserList(Userindex)
        If .flags.AdminInvisible = 0 Then
            ' Sacamos el mimetizmo
            If .flags.Mimetizado = 1 Then
                .Char.Body = .CharMimetizado.Body
                .Char.Head = .CharMimetizado.Head
                .Char.CascoAnim = .CharMimetizado.CascoAnim
                .Char.ShieldAnim = .CharMimetizado.ShieldAnim
                .Char.WeaponAnim = .CharMimetizado.WeaponAnim
                .Counters.Mimetismo = 0
                .flags.Mimetizado = 0
            End If
            
            .flags.AdminInvisible = 1
            .flags.invisible = 1
            .flags.Oculto = 1
            .flags.OldBody = .Char.Body
            .flags.OldHead = .Char.Head
            .Char.Body = 0
            .Char.Head = 0
            
            ' Solo el admin sabe que se hace invi
            Call EnviarDatosASlot(Userindex, PrepareMessageSetInvisible(.Char.CharIndex, True))
            'Le mandamos el mensaje para que borre el personaje a los clientes que estén cerca
            Call SendData(SendTarget.ToPCAreaButIndex, Userindex, PrepareMessageCharacterRemove(.Char.CharIndex))
        Else
            .flags.AdminInvisible = 0
            .flags.invisible = 0
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            .Char.Body = .flags.OldBody
            .Char.Head = .flags.OldHead
            
            'Borramos el personaje en del cliente del GM
            Call EnviarDatosASlot(Userindex, PrepareMessageCharacterRemove(.Char.CharIndex))
            'Le mandamos el mensaje para crear el personaje a los clientes que estén cerca
            Call MakeUserChar(True, .Pos.map, Userindex, .Pos.map, .Pos.X, .Pos.Y)
        End If
    End With
    
End Sub

Sub TratarDeHacerFogata(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Userindex As Integer)

Dim Suerte As Byte
Dim exito As Byte
Dim Obj As Obj
Dim posMadera As WorldPos

If Not LegalPos(map, X, Y) Then Exit Sub

With posMadera
    .map = map
    .X = X
    .Y = Y
End With

If MapData(map, X, Y).ObjInfo.ObjIndex <> 58 Then
    Call WriteConsoleMsg(Userindex, "Necesitas clickear sobre Leña para hacer ramitas", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If Distancia(posMadera, UserList(Userindex).Pos) > 2 Then
    Call WriteConsoleMsg(Userindex, "Estás demasiado lejos para prender la fogata.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If UserList(Userindex).flags.Muerto = 1 Then
    Call WriteConsoleMsg(Userindex, "No puedes hacer fogatas estando muerto.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If MapData(map, X, Y).ObjInfo.amount < 3 Then
    Call WriteConsoleMsg(Userindex, "Necesitas por lo menos tres troncos para hacer una fogata.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If


If UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) >= 0 And UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) < 6 Then
    Suerte = 3
ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) >= 6 And UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) <= 34 Then
    Suerte = 2
ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) >= 35 Then
    Suerte = 1
End If

exito = RandomNumber(1, Suerte)

If exito = 1 Then
    Obj.ObjIndex = FOGATA_APAG
    Obj.amount = MapData(map, X, Y).ObjInfo.amount \ 3
    
    Call WriteConsoleMsg(Userindex, "Has hecho " & Obj.amount & " fogatas.", FontTypeNames.FONTTYPE_INFO)
    
    Call MakeObj(Obj, map, X, Y)
    
    'Seteamos la fogata como el nuevo TargetObj del user
    UserList(Userindex).flags.TargetObj = FOGATA_APAG
Else
    '[CDT 17-02-2004]
    If Not UserList(Userindex).flags.UltimoMensaje = 10 Then
        Call WriteConsoleMsg(Userindex, "No has podido hacer la fogata.", FontTypeNames.FONTTYPE_INFO)
        UserList(Userindex).flags.UltimoMensaje = 10
    End If
    '[/CDT]
End If

Call SubirSkill(Userindex, Supervivencia)


End Sub

Public Sub DoPescar(ByVal Userindex As Integer)
On Error GoTo Errhandler

Dim Suerte As Integer
Dim res As Integer

If UserList(Userindex).clase = eClass.trabajador Then
    Call QuitarSta(Userindex, EsfuerzoPescarPescador)
Else
    Call QuitarSta(Userindex, EsfuerzoPescarGeneral)
End If

Dim Skill As Integer
Skill = UserList(Userindex).Stats.UserSkills(eSkill.Pesca)
Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)

res = RandomNumber(1, Suerte)

If res <= 6 Then
    Dim MiObj As Obj
    
    If UserList(Userindex).clase = eClass.trabajador Then
        MiObj.amount = RandomNumber(4, 12)
    Else
        MiObj.amount = 1
    End If
    
    'Nueva forma de extracción del pez [MaxTus]
    If MapInfo(UserList(Userindex).Pos.map).Pk = False Then
        If UserList(Userindex).flags.Navegando = False Then
            MiObj.ObjIndex = Pargo
        Else
            If res >= 3 Then
                MiObj.ObjIndex = PezEspada
            Else
                MiObj.ObjIndex = Lisa
            End If
        End If
    Else
        If UserList(Userindex).flags.Navegando = False Then
            MiObj.ObjIndex = Merluza
        Else
            MiObj.ObjIndex = Hipocampo
        End If
    End If
    
    If Not MeterItemEnInventario(Userindex, MiObj) Then
        Call TirarItemAlPiso(UserList(Userindex).Pos, MiObj)
    End If
    
    Call WriteConsoleMsg(Userindex, "¡Has pescado un lindo pez!", FontTypeNames.FONTTYPE_INFO)
    
Else
    '[CDT 17-02-2004]
    If Not UserList(Userindex).flags.UltimoMensaje = 6 Then
      Call WriteConsoleMsg(Userindex, "¡No has pescado Ninguno!", FontTypeNames.FONTTYPE_INFO)
      UserList(Userindex).flags.UltimoMensaje = 6
    End If
    '[/CDT]
End If

Call SubirSkill(Userindex, Pesca)

UserList(Userindex).Reputacion.PlebeRep = UserList(Userindex).Reputacion.PlebeRep + vlProleta
If UserList(Userindex).Reputacion.PlebeRep > MAXREP Then _
    UserList(Userindex).Reputacion.PlebeRep = MAXREP

UserList(Userindex).Counters.Trabajando = UserList(Userindex).Counters.Trabajando + 1

Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_PESCAR, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

Exit Sub

Errhandler:
    Call LogError("Error en DoPescar. Error " & Err.Number & " : " & Err.description)
End Sub

Public Sub DoPescarRed(ByVal Userindex As Integer)
On Error GoTo Errhandler

Dim iSkill As Integer
Dim Suerte As Integer
Dim res As Integer
Dim EsPescador As Boolean

If UserList(Userindex).clase = eClass.trabajador Then
    Call QuitarSta(Userindex, EsfuerzoPescarPescador)
    EsPescador = True
Else
    Call QuitarSta(Userindex, EsfuerzoPescarGeneral)
    EsPescador = False
End If

iSkill = UserList(Userindex).Stats.UserSkills(eSkill.Pesca)

' m = (60-11)/(1-10)
' y = mx - m*10 + 11

Suerte = Int(-0.00125 * iSkill * iSkill - 0.3 * iSkill + 49)

If Suerte > 0 Then
    res = RandomNumber(1, Suerte)
    
    If res < 6 Then
        Dim MiObj As Obj
        
        If EsPescador = True Then
            MiObj.amount = RandomNumber(8, 24)
        Else
            MiObj.amount = 3
        End If
        
        'Nueva forma de extracción del pez [MaxTus]
        If MapInfo(UserList(Userindex).Pos.map).Pk = False Then
            If UserList(Userindex).flags.Navegando = False Then
                MiObj.ObjIndex = Pargo
            Else
                If res >= 3 Then
                    MiObj.ObjIndex = PezEspada
                Else
                    MiObj.ObjIndex = Lisa
                End If
            End If
        Else
            If UserList(Userindex).flags.Navegando = False Then
                MiObj.ObjIndex = Merluza
            Else
                MiObj.ObjIndex = Hipocampo
            End If
        End If
        
        If Not MeterItemEnInventario(Userindex, MiObj) Then
            Call TirarItemAlPiso(UserList(Userindex).Pos, MiObj)
        End If
        
        Call WriteConsoleMsg(Userindex, "¡Has pescado algunos peces!", FontTypeNames.FONTTYPE_INFO)
        
    Else
        Call WriteConsoleMsg(Userindex, "¡No has pescado Ninguno!", FontTypeNames.FONTTYPE_INFO)
    End If
    
    Call SubirSkill(Userindex, Pesca)
End If

Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_PESCAR, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

    UserList(Userindex).Reputacion.PlebeRep = UserList(Userindex).Reputacion.PlebeRep + vlProleta
    If UserList(Userindex).Reputacion.PlebeRep > MAXREP Then _
        UserList(Userindex).Reputacion.PlebeRep = MAXREP
        
Exit Sub

Errhandler:
    Call LogError("Error en DoPescarRed")
End Sub

''
' Try to steal an item / gold to another character
'
' @param LadrOnIndex Specifies reference to user that stoles
' @param VictimaIndex Specifies reference to user that is being stolen

Public Sub DoRobar(ByVal LadrOnIndex As Integer, ByVal VictimaIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 24/07/028
'Last Modification By: Marco Vanotti (MarKoxX)
' - 24/07/08 Now it calls to WriteUpdateGold(VictimaIndex and LadrOnIndex) when the thief stoles gold. (MarKoxX)
'*************************************************

On Error GoTo Errhandler

If Not MapInfo(UserList(VictimaIndex).Pos.map).Pk Then Exit Sub

If UserList(LadrOnIndex).flags.Seguro Then
    Call WriteConsoleMsg(LadrOnIndex, "Debes quitar el seguro para robar", FontTypeNames.FONTTYPE_FIGHT)
    Exit Sub
End If

If TriggerZonaPelea(LadrOnIndex, VictimaIndex) <> TRIGGER6_AUSENTE Then Exit Sub

If UserList(VictimaIndex).Faccion.FuerzasCaos = 1 And UserList(LadrOnIndex).Faccion.FuerzasCaos = 1 Then
    Call WriteConsoleMsg(LadrOnIndex, "No puedes robar a otros miembros de las fuerzas del caos", FontTypeNames.FONTTYPE_FIGHT)
    Exit Sub
End If


Call QuitarSta(LadrOnIndex, 15)

Dim GuantesHurto As Boolean
'Tiene los Guantes de Hurto equipados?
GuantesHurto = True
If UserList(LadrOnIndex).Invent.AnilloEqpObjIndex = 0 Then
    GuantesHurto = False
Else
    If ObjData(UserList(LadrOnIndex).Invent.AnilloEqpObjIndex).DefensaMagicaMin <> 0 Then GuantesHurto = False
    If ObjData(UserList(LadrOnIndex).Invent.AnilloEqpObjIndex).DefensaMagicaMax <> 0 Then GuantesHurto = False
End If


If UserList(VictimaIndex).flags.Privilegios And PlayerType.User Then
    Dim Suerte As Integer
    Dim res As Integer
    
    If UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 10 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= -1 Then
                        Suerte = 35
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 20 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 11 Then
                        Suerte = 30
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 30 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 21 Then
                        Suerte = 28
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 40 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 31 Then
                        Suerte = 24
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 50 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 41 Then
                        Suerte = 22
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 60 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 51 Then
                        Suerte = 20
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 70 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 61 Then
                        Suerte = 18
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 80 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 71 Then
                        Suerte = 15
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 90 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 81 Then
                        Suerte = 10
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) < 100 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 91 Then
                        Suerte = 7
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) = 100 Then
                        Suerte = 5
    End If
    res = RandomNumber(1, Suerte)
        
    If res < 3 Then 'Exito robo
       
        If (RandomNumber(1, 50) < 25) And (UserList(LadrOnIndex).clase = eClass.Thief) Then
            If TieneObjetosRobables(VictimaIndex) Then
                Call RobarObjeto(LadrOnIndex, VictimaIndex)
            Else
                Call WriteConsoleMsg(LadrOnIndex, UserList(VictimaIndex).Name & " no tiene objetos.", FontTypeNames.FONTTYPE_INFO)
            End If
        Else 'Roba oro
            If UserList(VictimaIndex).Stats.GLD > 0 Then
                Dim N As Integer
                
                If UserList(LadrOnIndex).clase = eClass.Thief Then
                ' Si no tine puestos los guantes de hurto roba un 20% menos. Pablo (ToxicWaste)
                    If GuantesHurto Then
                        N = RandomNumber(100, 1000)
                    Else
                        N = RandomNumber(80, 800)
                    End If
                Else
                    N = RandomNumber(1, 100)
                End If
                If N > UserList(VictimaIndex).Stats.GLD Then N = UserList(VictimaIndex).Stats.GLD
                UserList(VictimaIndex).Stats.GLD = UserList(VictimaIndex).Stats.GLD - N
                
                UserList(LadrOnIndex).Stats.GLD = UserList(LadrOnIndex).Stats.GLD + N
                If UserList(LadrOnIndex).Stats.GLD > MAXORO Then _
                    UserList(LadrOnIndex).Stats.GLD = MAXORO
                
                Call WriteConsoleMsg(LadrOnIndex, "Le has robado " & N & " monedas de oro a " & UserList(VictimaIndex).Name, FontTypeNames.FONTTYPE_INFO)
                Call WriteUpdateGold(LadrOnIndex) 'Le actualizamos la billetera al ladron
                
                Call WriteUpdateGold(VictimaIndex) 'Le actualizamos la billetera a la victima
                Call FlushBuffer(VictimaIndex)
            Else
                Call WriteConsoleMsg(LadrOnIndex, UserList(VictimaIndex).Name & " no tiene oro.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
    Else
        Call WriteConsoleMsg(LadrOnIndex, "¡No has logrado robar Ninguno!", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(VictimaIndex, "¡" & UserList(LadrOnIndex).Name & " ha intentado robarte!", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(VictimaIndex, "¡" & UserList(LadrOnIndex).Name & " es un criminal!", FontTypeNames.FONTTYPE_INFO)
        Call FlushBuffer(VictimaIndex)
    End If

    If Not criminal(LadrOnIndex) Then
        Call VolverCriminal(LadrOnIndex)
    End If
    
    If UserList(LadrOnIndex).Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(LadrOnIndex)

    UserList(LadrOnIndex).Reputacion.LadronesRep = UserList(LadrOnIndex).Reputacion.LadronesRep + vlLadron
    If UserList(LadrOnIndex).Reputacion.LadronesRep > MAXREP Then _
        UserList(LadrOnIndex).Reputacion.LadronesRep = MAXREP
    Call SubirSkill(LadrOnIndex, Robar)
End If

Exit Sub

Errhandler:
    Call LogError("Error en DoRobar. Error " & Err.Number & " : " & Err.description)

End Sub

''
' Check if one item is stealable
'
' @param VictimaIndex Specifies reference to victim
' @param Slot Specifies reference to victim's inventory slot
' @return If the item is stealable
Public Function ObjEsRobable(ByVal VictimaIndex As Integer, ByVal Slot As Integer) As Boolean
' Agregué los barcos
' Esta funcion determina qué objetos son robables.

Dim OI As Integer

OI = UserList(VictimaIndex).Invent.Object(Slot).ObjIndex

ObjEsRobable = _
ObjData(OI).OBJType <> eOBJType.otLlaves And _
UserList(VictimaIndex).Invent.Object(Slot).Equipped = 0 And _
ObjData(OI).Real = 0 And _
ObjData(OI).Caos = 0 And _
ObjData(OI).Canjeable = 1 And _
ObjData(OI).OBJType <> eOBJType.otBarcos

End Function

''
' Try to steal an item to another character
'
' @param LadrOnIndex Specifies reference to user that stoles
' @param VictimaIndex Specifies reference to user that is being stolen
Public Sub RobarObjeto(ByVal LadrOnIndex As Integer, ByVal VictimaIndex As Integer)
'Call LogTarea("Sub RobarObjeto")
Dim flag As Boolean
Dim i As Integer
flag = False

If RandomNumber(1, 12) < 6 Then 'Comenzamos por el principio o el final?
    i = 1
    Do While Not flag And i <= MAX_INVENTORY_SLOTS
        'Hay objeto en este slot?
        If UserList(VictimaIndex).Invent.Object(i).ObjIndex > 0 Then
           If ObjEsRobable(VictimaIndex, i) Then
                 If RandomNumber(1, 10) < 4 Then flag = True
           End If
        End If
        If Not flag Then i = i + 1
    Loop
Else
    i = 20
    Do While Not flag And i > 0
      'Hay objeto en este slot?
      If UserList(VictimaIndex).Invent.Object(i).ObjIndex > 0 Then
         If ObjEsRobable(VictimaIndex, i) Then
               If RandomNumber(1, 10) < 4 Then flag = True
         End If
      End If
      If Not flag Then i = i - 1
    Loop
End If

If flag Then
    Dim MiObj As Obj
    Dim num As Byte
    'Cantidad al azar
    num = RandomNumber(1, 5)
                
    If num > UserList(VictimaIndex).Invent.Object(i).amount Then
         num = UserList(VictimaIndex).Invent.Object(i).amount
    End If
                
    MiObj.amount = num
    MiObj.ObjIndex = UserList(VictimaIndex).Invent.Object(i).ObjIndex
    
    UserList(VictimaIndex).Invent.Object(i).amount = UserList(VictimaIndex).Invent.Object(i).amount - num
                
    If UserList(VictimaIndex).Invent.Object(i).amount <= 0 Then
          Call QuitarUserInvItem(VictimaIndex, CByte(i), 1)
    End If
            
    Call UpdateUserInv(False, VictimaIndex, CByte(i))
                
    If Not MeterItemEnInventario(LadrOnIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(LadrOnIndex).Pos, MiObj)
    End If
    
    If UserList(LadrOnIndex).clase = eClass.Thief Then
        Call WriteConsoleMsg(LadrOnIndex, "Has robado " & MiObj.amount & " " & ObjData(MiObj.ObjIndex).Name, FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(LadrOnIndex, "Has hurtado " & MiObj.amount & " " & ObjData(MiObj.ObjIndex).Name, FontTypeNames.FONTTYPE_INFO)
    End If
Else
    Call WriteConsoleMsg(LadrOnIndex, "No has logrado robar ningún objeto.", FontTypeNames.FONTTYPE_INFO)
End If

'If exiting, cancel de quien es robado
Call CancelExit(VictimaIndex)

End Sub

Public Sub DoApuñalar(ByVal Userindex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal daño As Integer)
'***************************************************
'Autor: Nacho (Integer) & Unknown (orginal version)
'Last Modification: 04/17/08 - (NicoNZ)
'Simplifique la cuenta que hacia para sacar la suerte
'y arregle la cuenta que hacia para sacar el daño
'***************************************************
Dim Suerte As Integer
Dim Skill As Integer

Skill = UserList(Userindex).Stats.UserSkills(eSkill.Apuñalar)

Select Case UserList(Userindex).clase
    Case eClass.Assasin
        Suerte = Int(((0.00003 * Skill - 0.002) * Skill + 0.098) * Skill + 4.25)
    
    Case eClass.Cleric, eClass.Paladin
        Suerte = Int(((0.000003 * Skill + 0.0006) * Skill + 0.0107) * Skill + 4.93)
    
    Case eClass.Bard
        Suerte = Int(((0.000002 * Skill + 0.0002) * Skill + 0.032) * Skill + 4.81)
    
    Case Else
        Suerte = Int(0.0361 * Skill + 4.39)
End Select


If RandomNumber(0, 100) < Suerte Then
    If VictimUserIndex <> 0 Then
        If UserList(Userindex).clase = eClass.Assasin Then
            daño = Round(daño * 1.4, 0)
        Else
            daño = Round(daño * 1.5, 0)
        End If
        
        UserList(VictimUserIndex).Stats.MinHP = UserList(VictimUserIndex).Stats.MinHP - daño
        Call WriteConsoleMsg(Userindex, "Has apuñalado a " & UserList(VictimUserIndex).Name & " por " & daño, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(VictimUserIndex, "Te ha apuñalado " & UserList(Userindex).Name & " por " & daño, FontTypeNames.FONTTYPE_FIGHT)
        
        Call FlushBuffer(VictimUserIndex)
    Else
        Npclist(VictimNpcIndex).Stats.MinHP = Npclist(VictimNpcIndex).Stats.MinHP - Int(daño * 2)
        Call WriteConsoleMsg(Userindex, "Has apuñalado la criatura por " & Int(daño * 2), FontTypeNames.FONTTYPE_FIGHT)
        Call SubirSkill(Userindex, Apuñalar)
        '[Alejo]
        Call CalcularDarExp(Userindex, VictimNpcIndex, daño * 2)
    End If
Else
    Call WriteConsoleMsg(Userindex, "¡No has logrado apuñalar a tu enemigo!", FontTypeNames.FONTTYPE_FIGHT)
End If

End Sub

Public Sub DoGolpeCritico(ByVal Userindex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal daño As Integer)
'***************************************************
'Autor: Pablo (ToxicWaste)
'Last Modification: 28/01/2007
'***************************************************
Dim Suerte As Integer
Dim Skill As Integer

If UserList(Userindex).clase <> eClass.Bandit Then Exit Sub
If UserList(Userindex).Invent.WeaponEqpSlot = 0 Then Exit Sub
If ObjData(UserList(Userindex).Invent.WeaponEqpObjIndex).Name <> "Espada Vikinga" Then Exit Sub


Skill = UserList(Userindex).Stats.UserSkills(eSkill.Wrestling)

Suerte = Int((((0.00000003 * Skill + 0.000006) * Skill + 0.000107) * Skill + 0.0493) * 100)

If RandomNumber(0, 100) < Suerte Then
    daño = Int(daño * 0.5)
    If VictimUserIndex <> 0 Then
        UserList(VictimUserIndex).Stats.MinHP = UserList(VictimUserIndex).Stats.MinHP - daño
        Call WriteConsoleMsg(Userindex, "Has golpeado críticamente a " & UserList(VictimUserIndex).Name & " por " & daño, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(VictimUserIndex, UserList(Userindex).Name & " te ha golpeado críticamente por " & daño, FontTypeNames.FONTTYPE_FIGHT)
    Else
        Npclist(VictimNpcIndex).Stats.MinHP = Npclist(VictimNpcIndex).Stats.MinHP - daño
        Call WriteConsoleMsg(Userindex, "Has golpeado críticamente a la criatura por " & daño, FontTypeNames.FONTTYPE_FIGHT)
        '[Alejo]
        Call CalcularDarExp(Userindex, VictimNpcIndex, daño)
    End If
End If

End Sub

Public Sub QuitarSta(ByVal Userindex As Integer, ByVal Cantidad As Integer)

On Error GoTo Errhandler

    UserList(Userindex).Stats.MinSta = UserList(Userindex).Stats.MinSta - Cantidad
    If UserList(Userindex).Stats.MinSta < 0 Then UserList(Userindex).Stats.MinSta = 0
    Call WriteUpdateSta(Userindex)
    
Exit Sub

Errhandler:
    Call LogError("Error en QuitarSta. Error " & Err.Number & " : " & Err.description)
    
End Sub

Public Sub DoTalar(ByVal Userindex As Integer)
On Error GoTo Errhandler

Dim Suerte As Integer
Dim res As Integer

If UserList(Userindex).clase = eClass.trabajador Then
    Call QuitarSta(Userindex, EsfuerzoTalarLeñador)
Else
    Call QuitarSta(Userindex, EsfuerzoTalarGeneral)
End If

Dim Skill As Integer
Skill = UserList(Userindex).Stats.UserSkills(eSkill.Talar)
Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)

res = RandomNumber(1, Suerte)

If res <= 6 Then
    Dim MiObj As Obj
    
    If UserList(Userindex).clase = eClass.trabajador Then
        MiObj.amount = RandomNumber(4, 12)
    Else
        MiObj.amount = 1
    End If
    
    MiObj.ObjIndex = Leña
    
    
    If Not MeterItemEnInventario(Userindex, MiObj) Then
        
        Call TirarItemAlPiso(UserList(Userindex).Pos, MiObj)
        
    End If
    
    Call WriteConsoleMsg(Userindex, "¡Has conseguido algo de leña!", FontTypeNames.FONTTYPE_INFO)
    
Else
    '[CDT 17-02-2004]
    If Not UserList(Userindex).flags.UltimoMensaje = 8 Then
        Call WriteConsoleMsg(Userindex, "¡No has obtenido leña!", FontTypeNames.FONTTYPE_INFO)
        UserList(Userindex).flags.UltimoMensaje = 8
    End If
    '[/CDT]
End If

Call SubirSkill(Userindex, eSkill.Talar)

UserList(Userindex).Reputacion.PlebeRep = UserList(Userindex).Reputacion.PlebeRep + vlProleta
If UserList(Userindex).Reputacion.PlebeRep > MAXREP Then _
    UserList(Userindex).Reputacion.PlebeRep = MAXREP

UserList(Userindex).Counters.Trabajando = UserList(Userindex).Counters.Trabajando + 1

Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_TALAR, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

Exit Sub

Errhandler:
    Call LogError("Error en DoTalar")

End Sub
Public Sub DoMineria(ByVal Userindex As Integer)
On Error GoTo Errhandler

Dim Suerte As Integer
Dim res As Integer

If UserList(Userindex).clase = eClass.trabajador Then
    Call QuitarSta(Userindex, EsfuerzoExcavarMinero)
Else
    Call QuitarSta(Userindex, EsfuerzoExcavarGeneral)
End If

Dim Skill As Integer
Skill = UserList(Userindex).Stats.UserSkills(eSkill.Mineria)
Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)

res = RandomNumber(1, Suerte)

If res <= 5 Then
    Dim MiObj As Obj
    
    If UserList(Userindex).flags.TargetObj = 0 Then Exit Sub
    
    MiObj.ObjIndex = ObjData(UserList(Userindex).flags.TargetObj).MineralIndex
    
    If UserList(Userindex).clase = eClass.trabajador Then
        MiObj.amount = RandomNumber(6, 12) '(NicoNZ) 04/25/2008
    Else
        MiObj.amount = 1
    End If
    
    If Not MeterItemEnInventario(Userindex, MiObj) Then _
        Call TirarItemAlPiso(UserList(Userindex).Pos, MiObj)
    
    Call WriteConsoleMsg(Userindex, "¡Has extraido algunos minerales!", FontTypeNames.FONTTYPE_INFO)
    
Else
    '[CDT 17-02-2004]
    If Not UserList(Userindex).flags.UltimoMensaje = 9 Then
        Call WriteConsoleMsg(Userindex, "¡No has conseguido Ninguno!", FontTypeNames.FONTTYPE_INFO)
        UserList(Userindex).flags.UltimoMensaje = 9
    End If
    '[/CDT]
End If

Call SubirSkill(Userindex, Mineria)

UserList(Userindex).Reputacion.PlebeRep = UserList(Userindex).Reputacion.PlebeRep + vlProleta
If UserList(Userindex).Reputacion.PlebeRep > MAXREP Then _
    UserList(Userindex).Reputacion.PlebeRep = MAXREP

UserList(Userindex).Counters.Trabajando = UserList(Userindex).Counters.Trabajando + 1

Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_MINERO, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

Exit Sub

Errhandler:
    Call LogError("Error en Sub DoMineria")

End Sub

Public Sub DoMeditar(ByVal Userindex As Integer)

UserList(Userindex).Counters.IdleCount = 0

Dim Suerte As Integer
Dim res As Integer
Dim cant As Integer

'Barrin 3/10/03
'Esperamos a que se termine de concentrar
Dim TActual As Long
TActual = GetTickCount() And &H7FFFFFFF
If TActual - UserList(Userindex).Counters.tInicioMeditar < TIEMPO_INICIOMEDITAR Then
    Exit Sub
End If

If UserList(Userindex).Counters.bPuedeMeditar = False Then
    UserList(Userindex).Counters.bPuedeMeditar = True
End If
    
If UserList(Userindex).Stats.MinMAN >= UserList(Userindex).Stats.MaxMAN Then
    Call WriteConsoleMsg(Userindex, "Has terminado de meditar.", FontTypeNames.FONTTYPE_INFO)
    Call WriteMeditateToggle(Userindex)
    UserList(Userindex).flags.Meditando = False
     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateParticle(UserList(Userindex).Char.CharIndex, UserList(Userindex).Char.Particle, 1))
    UserList(Userindex).Char.Particle = 0
    UserList(Userindex).Char.loops = 0
    Exit Sub
End If

If UserList(Userindex).Stats.UserSkills(eSkill.Meditar) <= 10 _
   And UserList(Userindex).Stats.UserSkills(eSkill.Meditar) >= -1 Then
                    Suerte = 35
ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Meditar) <= 20 _
   And UserList(Userindex).Stats.UserSkills(eSkill.Meditar) >= 11 Then
                    Suerte = 30
ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Meditar) <= 30 _
   And UserList(Userindex).Stats.UserSkills(eSkill.Meditar) >= 21 Then
                    Suerte = 28
ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Meditar) <= 40 _
   And UserList(Userindex).Stats.UserSkills(eSkill.Meditar) >= 31 Then
                    Suerte = 24
ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Meditar) <= 50 _
   And UserList(Userindex).Stats.UserSkills(eSkill.Meditar) >= 41 Then
                    Suerte = 22
ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Meditar) <= 60 _
   And UserList(Userindex).Stats.UserSkills(eSkill.Meditar) >= 51 Then
                    Suerte = 20
ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Meditar) <= 70 _
   And UserList(Userindex).Stats.UserSkills(eSkill.Meditar) >= 61 Then
                    Suerte = 18
ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Meditar) <= 80 _
   And UserList(Userindex).Stats.UserSkills(eSkill.Meditar) >= 71 Then
                    Suerte = 15
ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Meditar) <= 90 _
   And UserList(Userindex).Stats.UserSkills(eSkill.Meditar) >= 81 Then
                    Suerte = 10
ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Meditar) < 100 _
   And UserList(Userindex).Stats.UserSkills(eSkill.Meditar) >= 91 Then
                    Suerte = 7
ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Meditar) = 100 Then
                    Suerte = 5
End If
res = RandomNumber(1, Suerte)

If res = 1 Then
    
    cant = Porcentaje(UserList(Userindex).Stats.MaxMAN, PorcentajeRecuperoMana)
    If cant <= 0 Then cant = 1
    UserList(Userindex).Stats.MinMAN = UserList(Userindex).Stats.MinMAN + cant
    If UserList(Userindex).Stats.MinMAN > UserList(Userindex).Stats.MaxMAN Then _
        UserList(Userindex).Stats.MinMAN = UserList(Userindex).Stats.MaxMAN
    
    Call WriteUpdateMana(Userindex)
    Call SubirSkill(Userindex, Meditar)
End If

End Sub

Public Sub DoHurtar(ByVal Userindex As Integer, ByVal VictimaIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modif: 28/01/2007
'Implements the pick pocket skill of the Bandit :)
'***************************************************
If UserList(Userindex).clase <> eClass.Bandit Then Exit Sub
'Esto es precario y feo, pero por ahora no se me ocurrió Ninguno mejor.
'Uso el slot de los anillos para "equipar" los guantes.
'Y los reconozco porque les puse DefensaMagicaMin y Max = 0
If UserList(Userindex).Invent.AnilloEqpObjIndex = 0 Then
    Exit Sub
Else
    If ObjData(UserList(Userindex).Invent.AnilloEqpObjIndex).DefensaMagicaMin <> 0 Then Exit Sub
    If ObjData(UserList(Userindex).Invent.AnilloEqpObjIndex).DefensaMagicaMax <> 0 Then Exit Sub
End If

Dim res As Integer
res = RandomNumber(1, 100)
If (res < 20) Then
    If TieneObjetosRobables(VictimaIndex) Then
        Call RobarObjeto(Userindex, VictimaIndex)
        Call WriteConsoleMsg(VictimaIndex, "¡" & UserList(Userindex).Name & " es un Bandido!", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(Userindex, UserList(VictimaIndex).Name & " no tiene objetos.", FontTypeNames.FONTTYPE_INFO)
    End If
End If

End Sub

Public Sub DoHandInmo(ByVal Userindex As Integer, ByVal VictimaIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modif: 17/02/2007
'Implements the special Skill of the Thief
'***************************************************
If UserList(VictimaIndex).flags.Paralizado = 1 Then Exit Sub
If UserList(Userindex).clase <> eClass.Thief Then Exit Sub
    
'una vez más, la forma de reconocer los guantes es medio patética.
If UserList(Userindex).Invent.AnilloEqpObjIndex = 0 Then
    Exit Sub
Else
    If ObjData(UserList(Userindex).Invent.AnilloEqpObjIndex).DefensaMagicaMin <> 0 Then Exit Sub
    If ObjData(UserList(Userindex).Invent.AnilloEqpObjIndex).DefensaMagicaMax <> 0 Then Exit Sub
End If

    
Dim res As Integer
res = RandomNumber(0, 100)
If res < (UserList(Userindex).Stats.UserSkills(eSkill.Wrestling) / 4) Then
    UserList(VictimaIndex).flags.Paralizado = 1
    UserList(VictimaIndex).Counters.Paralisis = IntervaloParalizado / 2
    Call WriteParalizeOK(VictimaIndex)
    Call WriteConsoleMsg(Userindex, "Tu golpe ha dejado inmóvil a tu oponente", FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(VictimaIndex, "¡El golpe te ha dejado inmóvil!", FontTypeNames.FONTTYPE_INFO)
End If

End Sub

Public Sub Desarmar(ByVal Userindex As Integer, ByVal VictimIndex As Integer)

Dim Suerte As Integer
Dim res As Integer

If UserList(Userindex).Stats.UserSkills(eSkill.Wrestling) <= 10 _
   And UserList(Userindex).Stats.UserSkills(eSkill.Wrestling) >= -1 Then
                    Suerte = 35
ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Wrestling) <= 20 _
   And UserList(Userindex).Stats.UserSkills(eSkill.Wrestling) >= 11 Then
                    Suerte = 30
ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Wrestling) <= 30 _
   And UserList(Userindex).Stats.UserSkills(eSkill.Wrestling) >= 21 Then
                    Suerte = 28
ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Wrestling) <= 40 _
   And UserList(Userindex).Stats.UserSkills(eSkill.Wrestling) >= 31 Then
                    Suerte = 24
ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Wrestling) <= 50 _
   And UserList(Userindex).Stats.UserSkills(eSkill.Wrestling) >= 41 Then
                    Suerte = 22
ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Wrestling) <= 60 _
   And UserList(Userindex).Stats.UserSkills(eSkill.Wrestling) >= 51 Then
                    Suerte = 20
ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Wrestling) <= 70 _
   And UserList(Userindex).Stats.UserSkills(eSkill.Wrestling) >= 61 Then
                    Suerte = 18
ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Wrestling) <= 80 _
   And UserList(Userindex).Stats.UserSkills(eSkill.Wrestling) >= 71 Then
                    Suerte = 15
ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Wrestling) <= 90 _
   And UserList(Userindex).Stats.UserSkills(eSkill.Wrestling) >= 81 Then
                    Suerte = 10
ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Wrestling) < 100 _
   And UserList(Userindex).Stats.UserSkills(eSkill.Wrestling) >= 91 Then
                    Suerte = 7
ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Wrestling) = 100 Then
                    Suerte = 5
End If
res = RandomNumber(1, Suerte)

If res <= 2 Then
        Call Desequipar(VictimIndex, UserList(VictimIndex).Invent.WeaponEqpSlot)
        Call WriteConsoleMsg(Userindex, "Has logrado desarmar a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
        If UserList(VictimIndex).Stats.ELV < 20 Then
            Call WriteConsoleMsg(VictimIndex, "Tu oponente te ha desarmado!", FontTypeNames.FONTTYPE_FIGHT)
        End If
        Call FlushBuffer(VictimIndex)
    End If
End Sub
