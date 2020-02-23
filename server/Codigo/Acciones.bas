Attribute VB_Name = "Acciones"
Option Explicit

''
' Modulo para manejar las acciones (doble click) de los carteles, foro, puerta, ramitas
'

''
' Ejecuta la accion del doble click
'
' @param UserIndex UserIndex
' @param Map Numero de mapa
' @param X X
' @param Y Y

Sub Accion(ByVal Userindex As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)
    Dim tempIndex As Integer
    
On Error Resume Next
    '¿Rango Visión? (ToxicWaste)
    If (Abs(UserList(Userindex).Pos.Y - Y) > RANGO_VISION_Y) Or (Abs(UserList(Userindex).Pos.X - X) > RANGO_VISION_X) Then
        Exit Sub
    End If
    
    '¿Posicion valida?
    If InMapBounds(map, X, Y) Then
        With UserList(Userindex)
            If MapData(map, X, Y).NpcIndex > 0 Then     'Acciones NPCs
                tempIndex = MapData(map, X, Y).NpcIndex
                
                'Set the target NPC
                .flags.TargetNPC = tempIndex
                
                If Npclist(tempIndex).Comercia = 1 Then
                    '¿Esta el user muerto? Si es asi no puede comerciar
                    If .flags.Muerto = 1 Then
                        Call WriteConsoleMsg(Userindex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'Is it already in commerce mode??
                    If .flags.Comerciando Then
                        Exit Sub
                    End If
                    
                    If Distancia(Npclist(tempIndex).Pos, .Pos) > 3 Then
                        Call WriteConsoleMsg(Userindex, "Estas demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'Iniciamos la rutina pa' comerciar.
                    Call IniciarComercioNPC(Userindex)
                
                ElseIf Npclist(tempIndex).NPCtype = eNPCType.Banquero Then
                    '¿Esta el user muerto? Si es asi no puede comerciar
                    If .flags.Muerto = 1 Then
                        Call WriteConsoleMsg(Userindex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'Is it already in commerce mode??
                    If .flags.Comerciando Then
                        Exit Sub
                    End If
                    
                    If Distancia(Npclist(tempIndex).Pos, .Pos) > 3 Then
                        Call WriteConsoleMsg(Userindex, "Estas demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'A depositar de una
                    Call IniciarDeposito(Userindex)
                
                ElseIf Npclist(tempIndex).NPCtype = eNPCType.Revividor Or Npclist(tempIndex).NPCtype = eNPCType.ResucitadorNewbie Then
                    If Distancia(.Pos, Npclist(tempIndex).Pos) > 10 Then
                        Call WriteConsoleMsg(Userindex, "El sacerdote no puede curarte debido a que estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'Revivimos si es necesario
                    If .flags.Muerto = 1 And (Npclist(tempIndex).NPCtype = eNPCType.Revividor Or EsNewbie(Userindex)) Then
                        Call RevivirUsuario(Userindex)
                    End If
                    
                    If Npclist(tempIndex).NPCtype = eNPCType.Revividor Or EsNewbie(Userindex) Then
                        'curamos totalmente
                        .Stats.MinHP = .Stats.MaxHP
                        Call WriteUpdateUserStats(Userindex)
                    End If
                ElseIf Npclist(tempIndex).NPCtype = eNPCType.Canjeros Then
                    '¿Esta el user muerto? Si es asi no puede comerciar
                    If .flags.Muerto = 1 Then
                        Call WriteConsoleMsg(Userindex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'Is it already in commerce mode??
                    If .flags.Comerciando Then
                        Exit Sub
                    End If
                    
                    If Distancia(Npclist(tempIndex).Pos, .Pos) > 3 Then
                        Call WriteConsoleMsg(Userindex, "Estas demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'A depositar de una
                    Call WriteCanjesInit(Userindex)
                End If
                
            '¿Es un obj?
            ElseIf MapData(map, X, Y).ObjInfo.ObjIndex > 0 Then
                tempIndex = MapData(map, X, Y).ObjInfo.ObjIndex
                
                .flags.TargetObj = tempIndex
                
                Select Case ObjData(tempIndex).OBJType
                    Case eOBJType.otPuertas 'Es una puerta
                        Call AccionParaPuerta(map, X, Y, Userindex)
                    Case eOBJType.otCarteles 'Es un cartel
                        Call AccionParaCartel(map, X, Y, Userindex)
                    Case eOBJType.otForos 'Foro
                        Call AccionParaForo(map, X, Y, Userindex)
                    Case eOBJType.otLeña    'Leña
                        If tempIndex = FOGATA_APAG And .flags.Muerto = 0 Then
                            Call AccionParaRamita(map, X, Y, Userindex)
                        End If
                End Select
            '>>>>>>>>>>>OBJETOS QUE OCUPAM MAS DE UN TILE<<<<<<<<<<<<<
            ElseIf MapData(map, X + 1, Y).ObjInfo.ObjIndex > 0 Then
                tempIndex = MapData(map, X + 1, Y).ObjInfo.ObjIndex
                .flags.TargetObj = tempIndex
                
                Select Case ObjData(tempIndex).OBJType
                    
                    Case eOBJType.otPuertas 'Es una puerta
                        Call AccionParaPuerta(map, X + 1, Y, Userindex)
                    
                End Select
            
            ElseIf MapData(map, X + 1, Y + 1).ObjInfo.ObjIndex > 0 Then
                tempIndex = MapData(map, X + 1, Y + 1).ObjInfo.ObjIndex
                .flags.TargetObj = tempIndex
        
                Select Case ObjData(tempIndex).OBJType
                    Case eOBJType.otPuertas 'Es una puerta
                        Call AccionParaPuerta(map, X + 1, Y + 1, Userindex)
                End Select
            
            ElseIf MapData(map, X, Y + 1).ObjInfo.ObjIndex > 0 Then
                tempIndex = MapData(map, X, Y + 1).ObjInfo.ObjIndex
                .flags.TargetObj = tempIndex
                
                Select Case ObjData(tempIndex).OBJType
                    Case eOBJType.otPuertas 'Es una puerta
                        Call AccionParaPuerta(map, X, Y + 1, Userindex)
                End Select
            End If
        End With
    End If
End Sub

Sub AccionParaForo(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Userindex As Integer)
On Error Resume Next

Dim Pos As WorldPos
Pos.map = map
Pos.X = X
Pos.Y = Y

If Distancia(Pos, UserList(Userindex).Pos) > 2 Then
    Call WriteConsoleMsg(Userindex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

'¿Hay mensajes?
Dim f As String, tit As String, men As String, BASE As String, auxcad As String
f = App.Path & "\foros\" & UCase$(ObjData(MapData(map, X, Y).ObjInfo.ObjIndex).ForoID) & ".for"
If FileExist(f, vbNormal) Then
    Dim num As Integer
    num = val(GetVar(f, "INFO", "CantMSG"))
    BASE = Left$(f, Len(f) - 4)
    Dim i As Integer
    Dim N As Integer
    For i = 1 To num
        N = FreeFile
        f = BASE & i & ".for"
        Open f For Input Shared As #N
        Input #N, tit
        men = vbNullString
        auxcad = vbNullString
        Do While Not EOF(N)
            Input #N, auxcad
            men = men & vbCrLf & auxcad
        Loop
        Close #N
        Call WriteAddForumMsg(Userindex, tit, men)
        
    Next
End If
Call WriteShowForumForm(Userindex)
End Sub


Sub AccionParaPuerta(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Userindex As Integer)
On Error Resume Next

If Not (Distance(UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y, X, Y) > 2) Then
    If ObjData(MapData(map, X, Y).ObjInfo.ObjIndex).Llave = 0 Then
        If ObjData(MapData(map, X, Y).ObjInfo.ObjIndex).Cerrada = 1 Then
                'Abre la puerta
                If ObjData(MapData(map, X, Y).ObjInfo.ObjIndex).Llave = 0 Then
                    
                    MapData(map, X, Y).ObjInfo.ObjIndex = ObjData(MapData(map, X, Y).ObjInfo.ObjIndex).IndexAbierta
                    
                    Call modSendData.SendToAreaByPos(map, X, Y, PrepareMessageObjectCreate(ObjData(MapData(map, X, Y).ObjInfo.ObjIndex).GrhIndex, X, Y))
                    
                    'Desbloquea
                    MapData(map, X, Y).Blocked = 0
                    MapData(map, X - 1, Y).Blocked = 0
                    
                    'Bloquea todos los mapas
                    Call Bloquear(True, map, X, Y, 0)
                    Call Bloquear(True, map, X - 1, Y, 0)
                    
                      
                    'Sonido
                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_PUERTA, X, Y))
                    
                Else
                     Call WriteConsoleMsg(Userindex, "La puerta esta cerrada con llave.", FontTypeNames.FONTTYPE_INFO)
                End If
        Else
                'Cierra puerta
                MapData(map, X, Y).ObjInfo.ObjIndex = ObjData(MapData(map, X, Y).ObjInfo.ObjIndex).IndexCerrada
                
                Call modSendData.SendToAreaByPos(map, X, Y, PrepareMessageObjectCreate(ObjData(MapData(map, X, Y).ObjInfo.ObjIndex).GrhIndex, X, Y))
                                
                MapData(map, X, Y).Blocked = 1
                MapData(map, X - 1, Y).Blocked = 1
                
                
                Call Bloquear(True, map, X - 1, Y, 1)
                Call Bloquear(True, map, X, Y, 1)
                
                Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_PUERTA, X, Y))
        End If
        
        UserList(Userindex).flags.TargetObj = MapData(map, X, Y).ObjInfo.ObjIndex
    Else
        Call WriteConsoleMsg(Userindex, "La puerta esta cerrada con llave.", FontTypeNames.FONTTYPE_INFO)
    End If
Else
    Call WriteConsoleMsg(Userindex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
End If

End Sub

Sub AccionParaCartel(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Userindex As Integer)
On Error Resume Next

If ObjData(MapData(map, X, Y).ObjInfo.ObjIndex).OBJType = 8 Then
  
  If Len(ObjData(MapData(map, X, Y).ObjInfo.ObjIndex).texto) > 0 Then
    Call WriteShowSignal(Userindex, MapData(map, X, Y).ObjInfo.ObjIndex)
  End If
  
End If

End Sub

Sub AccionParaRamita(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Userindex As Integer)
On Error Resume Next

Dim Suerte As Byte
Dim exito As Byte
Dim Obj As Obj

Dim Pos As WorldPos
Pos.map = map
Pos.X = X
Pos.Y = Y

If Distancia(Pos, UserList(Userindex).Pos) > 2 Then
    Call WriteConsoleMsg(Userindex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If MapData(map, X, Y).trigger = eTrigger.ZONASEGURA Or MapInfo(map).Pk = False Then
    Call WriteConsoleMsg(Userindex, "En zona segura no puedes hacer fogatas.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If UserList(Userindex).Stats.UserSkills(Supervivencia) > 1 And UserList(Userindex).Stats.UserSkills(Supervivencia) < 6 Then
            Suerte = 3
ElseIf UserList(Userindex).Stats.UserSkills(Supervivencia) >= 6 And UserList(Userindex).Stats.UserSkills(Supervivencia) <= 10 Then
            Suerte = 2
ElseIf UserList(Userindex).Stats.UserSkills(Supervivencia) >= 10 And UserList(Userindex).Stats.UserSkills(Supervivencia) Then
            Suerte = 1
End If

exito = RandomNumber(1, Suerte)

If exito = 1 Then
    If MapInfo(UserList(Userindex).Pos.map).Zona <> Ciudad Then
        Obj.ObjIndex = FOGATA
        Obj.amount = 1
        
        Call WriteConsoleMsg(Userindex, "Has prendido la fogata.", FontTypeNames.FONTTYPE_INFO)
        
        Call MakeObj(Obj, map, X, Y)
        
        'Las fogatas prendidas se deben eliminar
        Dim Fogatita As New cGarbage
        Fogatita.map = map
        Fogatita.X = X
        Fogatita.Y = Y
        Call TrashCollector.Add(Fogatita)
    Else
        Call WriteConsoleMsg(Userindex, "La ley impide realizar fogatas en las ciudades.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
Else
    Call WriteConsoleMsg(Userindex, "No has podido hacer fuego.", FontTypeNames.FONTTYPE_INFO)
End If

Call SubirSkill(Userindex, Supervivencia)

End Sub

Public Sub DoMetamorfosis(ByVal Userindex As Integer, Optional Body As Integer) 'Metamorfosis
    With UserList(Userindex)
    
        If Not .clase = 7 Then
            Call WriteConsoleMsg(Userindex, "Solo los druidas conocen el arte de la metamorfosis!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    
        If .flags.Muerto = 1 Then Exit Sub
        
        If .flags.Equitando = 1 Then
            Call WriteConsoleMsg(Userindex, "No puedes metamorfearte mientras equitas!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If .flags.Navegando = 1 Then
            Call WriteConsoleMsg(Userindex, "No puedes metamorfearte mientras navegas!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
         
        .Char.Head = 0
        .Char.Body = Body
        .Char.ShieldAnim = NingunEscudo
        .Char.WeaponAnim = NingunArma
        .Char.CascoAnim = NingunCasco
           
        .flags.Metamorfosis = 1
        
        Call ChangeUserChar(Userindex, .Char.Body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.Aura)
        'Call SendData(SendTarget.ToPCArea, Userindex, UserList(Userindex).Pos.map, "TW" & 17)
        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(.Char.CharIndex, FXIDs.FXWARP, 0))
    End With
End Sub
 
Public Sub EfectoMetamorfosis(ByVal Userindex As Integer) 'Metamorfosis
    With UserList(Userindex)
        If .Counters.Metamorfosis < IntervaloInvisible * 2 Then
            .Counters.Metamorfosis = .Counters.Metamorfosis + 1
        Else
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
         
            .flags.Metamorfosis = 0
            .Counters.Metamorfosis = 0
         
            Call ChangeUserChar(Userindex, .Char.Body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.Aura)
        End If
    End With
End Sub

