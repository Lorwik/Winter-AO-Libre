Attribute VB_Name = "modQuests"
'Amra
'Argentum Online 0.11.2.1
'Copyright (C) 2002 Márquez Pablo Ignacio
'
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
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez
'
'
'**************************************************************
' modQuests.bas - Realiza todos los handles para el sistema de
' Quests dentro del juego.
'
' Escrito y diseñado por Hernán Gurmendi a.k.a. Amraphen
' (hgurmen@hotmail.com)
'**************************************************************
Option Explicit

Public Type tQuest
    Nombre As String
    Descripcion As String
    NivelRequerido As Integer
    
    NpcKillIndex As Integer
    CantNPCs As Integer

    ObjIndex As Integer
    CantOBJs As Integer
    
    GLDReward As Long
    EXPReward As Long
    
    OBJRewardIndex As Integer
    CantOBJsReward As Integer
    
    Redoable As Byte
End Type

Public Type tUserQuest
    QuestIndex As Integer
    NPCsKilled As Integer
End Type

Public Const MAXUSERQUESTS As Byte = 10
Public QuestList() As tQuest

Public Sub LoadQuests()
'**************************************************************
'Author: Hernán Gurmendi (Amraphen)
'Last Modify Date: 13/10/2007
'Carga el archivo QUESTS.DAT.
'**************************************************************
Dim QuestFile As clsIniReader
Dim tmpInt As Integer

    Set QuestFile = New clsIniReader
    Call QuestFile.Initialize(App.Path & "\DAT\QUESTS.DAT")
       
    ReDim QuestList(1 To QuestFile.GetValue("INIT", "NumQuests"))
    
    For tmpInt = 1 To UBound(QuestList)
        QuestList(tmpInt).Nombre = QuestFile.GetValue("QUEST" & tmpInt, "Nombre")
        QuestList(tmpInt).Descripcion = QuestFile.GetValue("QUEST" & tmpInt, "Descripcion")
        
        QuestList(tmpInt).NpcKillIndex = QuestFile.GetValue("QUEST" & tmpInt, "NpcKillIndex")
        QuestList(tmpInt).CantNPCs = QuestFile.GetValue("QUEST" & tmpInt, "CantNPCs")
        
        QuestList(tmpInt).ObjIndex = QuestFile.GetValue("QUEST" & tmpInt, "OBJIndex")
        QuestList(tmpInt).CantOBJs = QuestFile.GetValue("QUEST" & tmpInt, "CantOBJs")
        
        QuestList(tmpInt).GLDReward = QuestFile.GetValue("QUEST" & tmpInt, "GLDReward")
        QuestList(tmpInt).EXPReward = QuestFile.GetValue("QUEST" & tmpInt, "EXPReward")
                
        QuestList(tmpInt).OBJRewardIndex = QuestFile.GetValue("QUEST" & tmpInt, "OBJRewardIndex")
        QuestList(tmpInt).CantOBJsReward = QuestFile.GetValue("QUEST" & tmpInt, "CantOBJsReward")
        QuestList(tmpInt).Redoable = QuestFile.GetValue("QUEST" & tmpInt, "Redoable")
    Next tmpInt
End Sub

Public Function UserTieneQuest(ByVal userindex As Integer, ByVal QuestNumber As Integer) As Integer
'**************************************************************
'Author: Hernán Gurmendi (Amraphen)
'Last Modify Date: 13/10/2007
'Devuelve 0 si no tiene la quest especificada en QuestNumber, o
'el numero de slot en el que tiene la quest.
'**************************************************************
Dim tmpInt As Integer

    For tmpInt = 1 To MAXUSERQUESTS
        If UserList(userindex).Stats.UserQuests(tmpInt).QuestIndex = QuestNumber Then
            UserTieneQuest = tmpInt
            Exit Function
        End If
    Next tmpInt
    
    UserTieneQuest = 0
End Function

Public Sub UserFinishQuest(ByVal userindex As Integer, ByVal QuestNumber As Integer)
'**************************************************************
'Author: Hernán Gurmendi (Amraphen)
'Last Modify Date: 13/10/2007
'Realiza el handle de /QUEST en caso de que el personaje ya
'tenga la quest.
'**************************************************************
Dim UTQ As Integer 'Determina el valor de UserTieneQuest
Dim tmpObj As Obj
Dim tmpInt As Integer

    If QuestList(QuestNumber).ObjIndex Then
        If TieneObjetos(QuestList(QuestNumber).ObjIndex, QuestList(QuestNumber).CantOBJs, userindex) = False Then
            Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & vbWhite & "°" & "Debes traerme los objetos que te he pedido antes de poder terminar la misión." & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
            Exit Sub
        End If
    End If
    
    UTQ = UserTieneQuest(userindex, QuestNumber)
    
    If QuestList(QuestNumber).NpcKillIndex Then
        If UserList(userindex).Stats.UserQuests(UTQ).NPCsKilled < QuestList(QuestNumber).CantNPCs Then
            Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & vbWhite & "°" & "Debes matar las criaturas que te he pedido antes de poder terminar la misión." & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
            Exit Sub
        End If
    End If
    
    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & vbWhite & "°" & "Gracias por ayudarme, noble aventurero, he aquí tu recompensa." & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
    Call SendData(SendTarget.Toindex, userindex, 0, "||Has completado la misión " & Chr(34) & QuestList(QuestNumber).Nombre & Chr(34) & "." & FONTTYPE_INFO)
                Call SonidosMapas.ReproducirSonido(SendTarget.Toindex, userindex, UserList(userindex).pos.Map, 240)
                Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & 17)
                
    If QuestList(QuestNumber).ObjIndex Then
        For tmpInt = 1 To MAX_INVENTORY_SLOTS
            If UserList(userindex).Invent.Object(tmpInt).ObjIndex = QuestList(QuestNumber).ObjIndex Then
                Call QuitarUserInvItem(userindex, CByte(tmpInt), QuestList(QuestNumber).CantOBJs)
                Exit For
            End If
        Next tmpInt
    End If
        
    If QuestList(QuestNumber).EXPReward Then
        UserList(userindex).Stats.Exp = UserList(userindex).Stats.Exp + QuestList(QuestNumber).EXPReward
        Call SendData(SendTarget.Toindex, userindex, 0, "||Has ganado " & QuestList(QuestNumber).EXPReward & " puntos de experiencia como recompensa." & FONTTYPE_INFO)
    End If
    
    If QuestList(QuestNumber).GLDReward Then
        UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD + QuestList(QuestNumber).GLDReward
        Call SendData(SendTarget.Toindex, userindex, 0, "||Has ganado " & QuestList(QuestNumber).GLDReward & " monedas de oro como recompensa." & FONTTYPE_INFO)
    End If
    
    If QuestList(QuestNumber).OBJRewardIndex Then
        tmpObj.ObjIndex = QuestList(QuestNumber).OBJRewardIndex
        tmpObj.Amount = QuestList(QuestNumber).CantOBJsReward
        
        If MeterItemEnInventario(userindex, tmpObj) = False Then
            Call TirarItemAlPiso(UserList(userindex).pos, tmpObj)
            Call SendData(SendTarget.Toindex, userindex, 0, "||Has recibido " & QuestList(QuestNumber).CantOBJsReward & " " & ObjData(QuestList(QuestNumber).OBJRewardIndex).name & " como recompensa." & FONTTYPE_INFO)
        End If
    End If
    
    UserList(userindex).Stats.UserQuests(UTQ).QuestIndex = 0
    UserList(userindex).Stats.UserQuests(UTQ).NPCsKilled = 0
    UserList(userindex).Stats.UserQuestsDone = UserList(userindex).Stats.UserQuestsDone & QuestNumber & "-"
    
    Call UpdateUserInv(True, userindex, 0)
    Call CheckUserLevel(userindex)
    Call SendUserStatsBox(userindex)
End Sub

Public Sub UserAceptarQuest(ByVal userindex As Integer, ByVal QuestNumber As Integer)
'**************************************************************
'Author: Hernán Gurmendi (Amraphen)
'Last Modify Date: 13/10/2007
'Realiza el handle de /QUEST en caso de que el personaje no
'tenga la quest.
'**************************************************************
Dim UFQS As Integer

    UFQS = UserFreeQuestSlot(userindex)
    
    If QuestList(QuestNumber).Redoable = 0 Then
        If UserHizoQuest(userindex, QuestNumber) = True Then
            Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & vbWhite & "°" & "Ya has hecho la misión." & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
            Exit Sub
        End If
    End If
    
    If UFQS = 0 Then
        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & vbWhite & "°" & "Debes terminar o cancelar alguna misión antes de poder aceptar otra." & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
        Exit Sub
    End If

    If UserList(userindex).Stats.ELV < QuestList(QuestNumber).NivelRequerido Then
        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & vbWhite & "°" & "No tienes nivel suficiente como para empezar esta misión." & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
        Exit Sub
    End If
    
    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & vbWhite & "°" & Npclist(UserList(userindex).flags.TargetNPC).TalkDuringQuest & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
    Call SendData(SendTarget.Toindex, userindex, 0, "||Has aceptado la misión " & Chr(34) & QuestList(QuestNumber).Nombre & Chr(34) & "." & FONTTYPE_INFO)
    Call SonidosMapas.ReproducirSonido(SendTarget.Toindex, userindex, UserList(userindex).pos.Map, 157)
     
    UserList(userindex).Stats.UserQuests(UFQS).QuestIndex = QuestNumber
    UserList(userindex).Stats.UserQuests(UFQS).NPCsKilled = 0
End Sub

Public Function UserFreeQuestSlot(ByVal userindex As Integer) As Integer
'**************************************************************
'Author: Hernán Gurmendi (Amraphen)
'Last Modify Date: 13/10/2007
'Devuelve 0 si no tiene algun slot de quest libre, o el primer
'slot de quest que tiene libre.
'**************************************************************
Dim tmpInt As Integer

    For tmpInt = 1 To MAXUSERQUESTS
        If UserList(userindex).Stats.UserQuests(tmpInt).QuestIndex = 0 Then
            UserFreeQuestSlot = tmpInt
            Exit Function
        End If
    Next tmpInt
    
    UserFreeQuestSlot = 0
End Function

Public Function UserHizoQuest(ByVal userindex As Integer, ByVal QuestNumber As Integer) As Boolean
'**************************************************************
'Author: Hernán Gurmendi (Amraphen)
'Last Modify Date: 13/10/2007
'Devuelve verdadero si el user hizo la quest QuestNumber, o
'falso si el user no la hizo.
'**************************************************************
Dim arrStr() As String
Dim tmpInt As Integer

    arrStr = Split(UserList(userindex).Stats.UserQuestsDone, "-")
    
    For tmpInt = 0 To UBound(arrStr)
        If CInt(arrStr(tmpInt)) = QuestNumber Then
            UserHizoQuest = True
            Exit Function
        End If
    Next tmpInt
    
    UserHizoQuest = False
End Function

Public Sub HandleQuest(ByVal userindex As Integer)
'**************************************************************
'Author: Hernán Gurmendi (Amraphen)
'Last Modify Date: 13/10/2007
'Realiza el handle del comando /QUEST.
'**************************************************************
Dim UTQ As Integer 'Determina el valor de la función UserTieneQuest.
Dim QN As Integer 'Determina el valor de la quest que posee el NPC.

    If Distancia(UserList(userindex).pos, Npclist(UserList(userindex).flags.TargetNPC).pos) > 4 Then
        Call SendData(SendTarget.Toindex, userindex, 0, "||No puedes hablar con el NPC ya que estas demasiado lejos." & FONTTYPE_INFO)
        Exit Sub
    End If

    If UserList(userindex).flags.TargetNPC = 0 Then
        Call SendData(SendTarget.Toindex, userindex, 0, "||Debes seleccionar un NPC con el cual hablar." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    If UserList(userindex).flags.Muerto Then
        Call SendData(SendTarget.Toindex, userindex, 0, "||Estás muerto!" & FONTTYPE_INFO)
        Exit Sub
    End If
    
    QN = Npclist(UserList(userindex).flags.TargetNPC).QuestNumber
    
    If QN = 0 Then
        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & vbWhite & "°" & "No tengo ninguna misión para tí." & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
        Exit Sub
    End If
    
    UTQ = UserTieneQuest(userindex, QN)
        
    If UTQ Then
        Call UserFinishQuest(userindex, QN)
    Else
        Call UserAceptarQuest(userindex, QN)
    End If
End Sub

Public Sub SendQuestList(ByVal userindex As Integer)
'**************************************************************
'Author: Hernán Gurmendi (Amraphen)
'Last Modify Date: 23/10/2007
'Envía a UserIndex la lista de quests.
'**************************************************************
Dim tmpString As String
Dim i As Integer

    For i = 1 To MAXUSERQUESTS
        If UserList(userindex).Stats.UserQuests(i).QuestIndex = 0 Then
            tmpString = tmpString & "0-"
        Else
            tmpString = tmpString & QuestList(UserList(userindex).Stats.UserQuests(i).QuestIndex).Nombre & "-"
        End If
    Next i
    
    Call SendData(SendTarget.Toindex, userindex, 0, "QL" & Left$(tmpString, Len(tmpString) - 1))
End Sub
'/Amra
