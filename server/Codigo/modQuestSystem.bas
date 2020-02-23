Attribute VB_Name = "modQuestSystem"
Option Explicit

'Constantes de las quests
Public Const MAXUSERQUESTS As Integer = 15     'Máxima cantidad de quests que puede tener un usuario al mismo tiempo.

Public Function TieneQuest(ByVal UserIndex As Integer, ByVal QuestNumber As Integer) As Byte
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Devuelve el slot de UserQuests en que tiene la quest QuestNumber. En caso contrario devuelve 0.
'Last modified: 27/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Dim i As Integer

    For i = 1 To MAXUSERQUESTS
        If UserList(UserIndex).QuestStats.Quests(i).QuestIndex = QuestNumber Then
            TieneQuest = i
            Exit Function
        End If
    Next i
    
    TieneQuest = 0
End Function

Public Function FreeQuestSlot(ByVal UserIndex As Integer) As Byte
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Devuelve el próximo slot de quest libre.
'Last modified: 27/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Dim i As Integer

    For i = 1 To MAXUSERQUESTS
        If UserList(UserIndex).QuestStats.Quests(i).QuestIndex = 0 Then
            FreeQuestSlot = i
            Exit Function
        End If
    Next i
    
    FreeQuestSlot = 0
End Function

Public Sub HandleQuestAccept(ByVal UserIndex As Integer)
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Maneja el evento de aceptar una quest.
'Last modified: 31/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Dim NpcIndex As Integer
Dim QuestSlot As Byte

    Call UserList(UserIndex).incomingData.ReadByte

    NpcIndex = UserList(UserIndex).flags.TargetNPC
    
    If NpcIndex = 0 Then Exit Sub
    
    'Está el personaje en la distancia correcta?
    If Distancia(UserList(UserIndex).Pos, Npclist(NpcIndex).Pos) > 5 Then
        Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    QuestSlot = FreeQuestSlot(UserIndex)
    
    'Agregamos la quest.
    With UserList(UserIndex).QuestStats.Quests(QuestSlot)
        .QuestIndex = Npclist(NpcIndex).QuestNumber
        
        If QuestList(.QuestIndex).RequiredNPCs Then ReDim .NPCsKilled(1 To QuestList(.QuestIndex).RequiredNPCs)
        Call WriteConsoleMsg(UserIndex, "Has aceptado la misión " & Chr(34) & QuestList(.QuestIndex).Nombre & Chr(34) & ".", FontTypeNames.FONTTYPE_INFO)
        
    End With
End Sub

Public Sub FinishQuest(ByVal UserIndex As Integer, ByVal QuestIndex As Integer, ByVal QuestSlot As Byte)
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Maneja el evento de terminar una quest.
'Last modified: 29/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Dim i As Integer
Dim InvSlotsLibres As Byte
Dim NpcIndex As Integer

    NpcIndex = UserList(UserIndex).flags.TargetNPC
    
    With QuestList(QuestIndex)
        'Comprobamos que tenga los objetos.
        If .RequiredOBJs > 0 Then
            For i = 1 To .RequiredOBJs
                If TieneObjetos(.RequiredOBJ(i).ObjIndex, .RequiredOBJ(i).amount, UserIndex) = False Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No has conseguido todos los objetos que te he pedido.", Npclist(NpcIndex).Char.CharIndex, vbWhite))
                    Exit Sub
                End If
            Next i
        End If
        
        'Comprobamos que haya matado todas las criaturas.
        If .RequiredNPCs > 0 Then
            For i = 1 To .RequiredNPCs
                If .RequiredNPC(i).amount > UserList(UserIndex).QuestStats.Quests(QuestSlot).NPCsKilled(i) Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No has matado todas las criaturas que te he pedido.", Npclist(NpcIndex).Char.CharIndex, vbWhite))
                    Exit Sub
                End If
            Next i
        End If
    
        'Comprobamos que el usuario tenga espacio para recibir los items.
        If .RewardOBJs > 0 Then
            'Buscamos la cantidad de slots de inventario libres.
            For i = 1 To MAX_INVENTORY_SLOTS
                If UserList(UserIndex).Invent.Object(i).ObjIndex = 0 Then InvSlotsLibres = InvSlotsLibres + 1
            Next i
            
            'Nos fijamos si entra
            If InvSlotsLibres < .RewardOBJs Then
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No tienes suficiente espacio en el inventario para recibir la recompensa. Vuelve cuando hayas hecho más espacio.", Npclist(NpcIndex).Char.CharIndex, vbWhite))
                Exit Sub
            End If
        End If
    
        'A esta altura ya cumplió los objetivos, entonces se le entregan las recompensas.
        Call WriteConsoleMsg(UserIndex, "¡Has completado la misión " & Chr(34) & QuestList(QuestIndex).Nombre & Chr(34) & "!", FontTypeNames.FONTTYPE_INFO)

        'Si la quest pedía objetos, se los saca al personaje.
        If .RequiredOBJs Then
            For i = 1 To .RequiredOBJs
                Call QuitarObjetos(.RequiredOBJ(i).ObjIndex, .RequiredOBJ(i).amount, UserIndex)
            Next i
        End If
        
        'Se entrega la experiencia.
        If .RewardEXP Then
            UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + .RewardEXP
            Call WriteConsoleMsg(UserIndex, "Has ganado " & .RewardEXP & " puntos de experiencia como recompensa.", FontTypeNames.FONTTYPE_EXP)
        End If
        
        'Se entrega el oro.
        If .RewardGLD Then
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + .RewardGLD
            Call WriteConsoleMsg(UserIndex, "Has ganado " & .RewardGLD & " monedas de oro como recompensa.", FontTypeNames.FONTTYPE_GLD)
        End If
        
        'Si hay recompensa de objetos, se entregan.
        If .RewardOBJs > 0 Then
            For i = 1 To .RewardOBJs
                If .RewardOBJ(i).amount Then
                    Call MeterItemEnInventario(UserIndex, .RewardOBJ(i))
                    Call WriteConsoleMsg(UserIndex, "Has recibido " & QuestList(QuestIndex).RewardOBJ(i).amount & " " & ObjData(QuestList(QuestIndex).RewardOBJ(i).ObjIndex).Name & " como recompensa.", FontTypeNames.FONTTYPE_INFO)
                End If
            Next i
        End If
    
        'Actualizamos el personaje
        Call CheckUserLevel(UserIndex)
        Call UpdateUserInv(True, UserIndex, 0)
    
        'Limpiamos el slot de quest.
        Call CleanQuestSlot(UserIndex, QuestSlot)
        
        'Ordenamos las quests
        Call ArrangeUserQuests(UserIndex)
    
        'Se agrega que el usuario ya hizo esta quest.
        Call AddDoneQuest(UserIndex, QuestIndex)
    End With
End Sub

Public Sub AddDoneQuest(ByVal UserIndex As Integer, ByVal QuestIndex As Integer)
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Agrega la quest QuestIndex a la lista de quests hechas.
'Last modified: 28/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    With UserList(UserIndex).QuestStats
        .NumQuestsDone = .NumQuestsDone + 1
        ReDim Preserve .QuestsDone(1 To .NumQuestsDone)
        .QuestsDone(.NumQuestsDone) = QuestIndex
    End With
End Sub

Public Function UserDoneQuest(ByVal UserIndex As Integer, ByVal QuestIndex As Integer) As Boolean
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Verifica si el usuario hizo la quest QuestIndex.
'Last modified: 28/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Dim i As Integer
    With UserList(UserIndex).QuestStats
        If .NumQuestsDone Then
            For i = 1 To .NumQuestsDone
                If .QuestsDone(i) = QuestIndex Then
                    UserDoneQuest = True
                    Exit Function
                End If
            Next i
        End If
    End With
    
    UserDoneQuest = False
        
End Function

Public Sub CleanQuestSlot(ByVal UserIndex As Integer, ByVal QuestSlot As Integer)
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Limpia un slot de quest de un usuario.
'Last modified: 28/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Dim i As Integer

    With UserList(UserIndex).QuestStats.Quests(QuestSlot)
        If .QuestIndex Then
            If QuestList(.QuestIndex).RequiredNPCs Then
                For i = 1 To QuestList(.QuestIndex).RequiredNPCs
                    .NPCsKilled(i) = 0
                Next i
            End If
        End If
        .QuestIndex = 0
    End With
End Sub

Public Sub ResetQuestStats(ByVal UserIndex As Integer)
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Limpia todos los QuestStats de un usuario
'Last modified: 28/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Dim i As Integer

    For i = 1 To MAXUSERQUESTS
        Call CleanQuestSlot(UserIndex, i)
    Next i
    
    With UserList(UserIndex).QuestStats
        .NumQuestsDone = 0
        Erase .QuestsDone
    End With
End Sub

Public Sub HandleQuest(ByVal UserIndex As Integer)
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Maneja el paquete Quest.
'Last modified: 28/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Dim NpcIndex As Integer
Dim tmpByte As Byte

    'Leemos el paquete
    Call UserList(UserIndex).incomingData.ReadByte

    NpcIndex = UserList(UserIndex).flags.TargetNPC
    
    If NpcIndex = 0 Then Exit Sub
    
    'Está el personaje en la distancia correcta?
    If Distancia(UserList(UserIndex).Pos, Npclist(NpcIndex).Pos) > 5 Then
        Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    'El NPC hace quests?
    If Npclist(NpcIndex).QuestNumber = 0 Then
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No tengo ninguna misión para ti.", Npclist(NpcIndex).Char.CharIndex, vbWhite))
        Exit Sub
    End If
    
    'El personaje ya hizo la quest?
    If UserDoneQuest(UserIndex, Npclist(NpcIndex).QuestNumber) Then
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("Ya has hecho una misión para mi.", Npclist(NpcIndex).Char.CharIndex, vbWhite))
        Exit Sub
    End If

    'El personaje tiene suficiente nivel?
    If UserList(UserIndex).Stats.ELV < QuestList(Npclist(NpcIndex).QuestNumber).RequiredLevel Then
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("Debes ser por lo menos nivel " & QuestList(Npclist(NpcIndex).QuestNumber).RequiredLevel & " para emprender esta misión.", Npclist(NpcIndex).Char.CharIndex, vbWhite))
        Exit Sub
    End If
    
    'A esta altura ya analizo todas las restricciones y esta preparado para el handle propiamente dicho

    tmpByte = TieneQuest(UserIndex, Npclist(NpcIndex).QuestNumber)
    
    If tmpByte Then
        'El usuario está haciendo la quest, entonces va a hablar con el NPC para recibir la recompensa.
        Call FinishQuest(UserIndex, Npclist(NpcIndex).QuestNumber, tmpByte)
    Else
        'El usuario no está haciendo la quest, entonces primero recibe un informe con los detalles de la misión.
        tmpByte = FreeQuestSlot(UserIndex)
        
        'El personaje tiene algun slot de quest para la nueva quest?
        If tmpByte = 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("Estás haciendo demasiadas misiones. Vuelve cuando hayas completado alguna.", Npclist(NpcIndex).Char.CharIndex, vbWhite))
            Exit Sub
        End If
        
        'Enviamos los detalles de la quest
        Call WriteQuestDetails(UserIndex, Npclist(NpcIndex).QuestNumber)
    End If
End Sub

Public Sub LoadQuests()
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Carga el archivo QUESTS.DAT en el array QuestList.
'Last modified: 27/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
On Error GoTo ErrorHandler
Dim Reader As clsIniReader
Dim NumQuests As Integer
Dim tmpStr As String
Dim i As Integer
Dim j As Integer
    
    'Cargamos el clsIniReader en memoria
    Set Reader = New clsIniReader
    
    'Lo inicializamos para el archivo QUESTS.DAT
    Call Reader.Initialize(DatPath & "QUESTS.DAT")
    
    'Redimensionamos el array
    NumQuests = Reader.GetValue("INIT", "NumQuests")
    ReDim QuestList(1 To NumQuests)
    
    'Cargamos los datos
    For i = 1 To NumQuests
        With QuestList(i)
            .Nombre = Reader.GetValue("QUEST" & i, "Nombre")
            .desc = Reader.GetValue("QUEST" & i, "Desc")
            .RequiredLevel = val(Reader.GetValue("QUEST" & i, "RequiredLevel"))
            
            'CARGAMOS OBJETOS REQUERIDOS
            .RequiredOBJs = val(Reader.GetValue("QUEST" & i, "RequiredOBJs"))
            If .RequiredOBJs > 0 Then
                ReDim .RequiredOBJ(1 To .RequiredOBJs)
                For j = 1 To .RequiredOBJs
                    tmpStr = Reader.GetValue("QUEST" & i, "RequiredOBJ" & j)
                    
                    .RequiredOBJ(j).ObjIndex = val(ReadField(1, tmpStr, 45))
                    .RequiredOBJ(j).amount = val(ReadField(2, tmpStr, 45))
                Next j
            End If
            
            'CARGAMOS NPCS REQUERIDOS
            .RequiredNPCs = val(Reader.GetValue("QUEST" & i, "RequiredNPCs"))
            If .RequiredNPCs > 0 Then
                ReDim .RequiredNPC(1 To .RequiredNPCs)
                For j = 1 To .RequiredNPCs
                    tmpStr = Reader.GetValue("QUEST" & i, "RequiredNPC" & j)
                    
                    .RequiredNPC(j).NpcIndex = val(ReadField(1, tmpStr, 45))
                    .RequiredNPC(j).amount = val(ReadField(2, tmpStr, 45))
                Next j
            End If
            
            .RewardGLD = val(Reader.GetValue("QUEST" & i, "RewardGLD"))
            .RewardEXP = val(Reader.GetValue("QUEST" & i, "RewardEXP"))
            
            'CARGAMOS OBJETOS DE RECOMPENSA
            .RewardOBJs = val(Reader.GetValue("QUEST" & i, "RewardOBJs"))
            If .RewardOBJs > 0 Then
                ReDim .RewardOBJ(1 To .RewardOBJs)
                For j = 1 To .RewardOBJs
                    tmpStr = Reader.GetValue("QUEST" & i, "RewardOBJ" & j)
                    
                    .RewardOBJ(j).ObjIndex = val(ReadField(1, tmpStr, 45))
                    .RewardOBJ(j).amount = val(ReadField(2, tmpStr, 45))
                Next j
            End If
        End With
    Next i
    
    'Eliminamos la clase
    Set Reader = Nothing
Exit Sub
                    
ErrorHandler:
    MsgBox "Error cargando el archivo QUESTS.DAT.", vbOKOnly + vbCritical
End Sub

Public Sub LoadQuestStats(ByVal UserIndex As Integer, ByRef UserFile As clsIniReader)
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Carga las QuestStats del usuario.
'Last modified: 28/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Dim i As Integer
Dim j As Integer
Dim tmpStr As String

    For i = 1 To MAXUSERQUESTS
        With UserList(UserIndex).QuestStats.Quests(i)
            tmpStr = UserFile.GetValue("QUESTS", "Q" & i)
            
            .QuestIndex = val(ReadField(1, tmpStr, 45))
            If .QuestIndex Then
                If QuestList(.QuestIndex).RequiredNPCs Then
                    ReDim .NPCsKilled(1 To QuestList(.QuestIndex).RequiredNPCs)
                    
                    For j = 1 To QuestList(.QuestIndex).RequiredNPCs
                        .NPCsKilled(j) = val(ReadField(j + 1, tmpStr, 45))
                    Next j
                End If
            End If
        End With
    Next i
    
    With UserList(UserIndex).QuestStats
        tmpStr = UserFile.GetValue("QUESTS", "QuestsDone")
        
        .NumQuestsDone = val(ReadField(1, tmpStr, 45))
        
        If .NumQuestsDone Then
            ReDim .QuestsDone(1 To .NumQuestsDone)
            For i = 1 To .NumQuestsDone
                .QuestsDone(i) = val(ReadField(i + 1, tmpStr, 45))
            Next i
        End If
    End With
                   
End Sub

Public Sub SaveQuestStats(ByVal UserIndex As Integer, ByVal UserFile As String)
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Guarda las QuestStats del usuario.
'Last modified: 29/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Dim i As Integer
Dim j As Integer
Dim tmpStr As String

    For i = 1 To MAXUSERQUESTS
        With UserList(UserIndex).QuestStats.Quests(i)
            tmpStr = .QuestIndex
            
            If .QuestIndex Then
                If QuestList(.QuestIndex).RequiredNPCs Then
                    For j = 1 To QuestList(.QuestIndex).RequiredNPCs
                        tmpStr = tmpStr & "-" & .NPCsKilled(j)
                    Next j
                End If
            End If
        
            Call WriteVar(UserFile, "QUESTS", "Q" & i, tmpStr)
        End With
    Next i
    
    With UserList(UserIndex).QuestStats
        tmpStr = .NumQuestsDone
        
        If .NumQuestsDone Then
            For i = 1 To .NumQuestsDone
                tmpStr = tmpStr & "-" & .QuestsDone(i)
            Next i
        End If
        
        Call WriteVar(UserFile, "QUESTS", "QuestsDone", tmpStr)
    End With
End Sub

Public Sub HandleQuestListRequest(ByVal UserIndex As Integer)
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Maneja el paquete QuestListRequest.
'Last modified: 30/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

    'Leemos el paquete
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call WriteQuestListSend(UserIndex)
End Sub

Public Sub ArrangeUserQuests(ByVal UserIndex As Integer)
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Ordena las quests del usuario de manera que queden todas al principio del arreglo.
'Last modified: 30/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Dim i As Integer
Dim j As Integer

    With UserList(UserIndex).QuestStats
        For i = 1 To MAXUSERQUESTS - 1
            If .Quests(i).QuestIndex = 0 Then
                For j = i + 1 To MAXUSERQUESTS
                    If .Quests(j).QuestIndex Then
                        .Quests(i) = .Quests(j)
                        Call CleanQuestSlot(UserIndex, j)
                        Exit For
                    End If
                Next j
            End If
        Next i
    End With
End Sub

Public Sub HandleQuestDetailsRequest(ByVal UserIndex As Integer)
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Maneja el paquete QuestInfoRequest.
'Last modified: 30/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Dim QuestSlot As Byte

    'Leemos el paquete
    Call UserList(UserIndex).incomingData.ReadByte
    
    QuestSlot = UserList(UserIndex).incomingData.ReadByte
    
    Call WriteQuestDetails(UserIndex, UserList(UserIndex).QuestStats.Quests(QuestSlot).QuestIndex, QuestSlot)
End Sub

Public Sub HandleQuestAbandon(ByVal UserIndex As Integer)
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Maneja el paquete QuestAbandon.
'Last modified: 31/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Leemos el paquete.
    Call UserList(UserIndex).incomingData.ReadByte
    
    'Borramos la quest.
    Call CleanQuestSlot(UserIndex, UserList(UserIndex).incomingData.ReadByte)
    
    'Ordenamos la lista de quests del usuario.
    Call ArrangeUserQuests(UserIndex)
    
    'Enviamos la lista de quests actualizada.
    Call WriteQuestListSend(UserIndex)
End Sub
