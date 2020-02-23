Attribute VB_Name = "modEdicion"
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
' modEdicion
'
' @remarks Funciones de Edicion
' @author gshaxor@gmail.com
' @version 0.1.38
' @date 20061016

Option Explicit

''
' Vacia el Deshacer
'
Public Sub Deshacer_Clear()
'*************************************************
'Author: ^[GS]^
'Last modified: 15/10/06
'*************************************************
Dim i As Integer
' Vacio todos los campos afectados
For i = 1 To maxDeshacer
    MapData_Deshacer_Info(i).Libre = True
Next
' no ahi que deshacer
frmMain.mnuDeshacer.Enabled = False
End Sub

''
' Agrega un Deshacer
'
Public Sub Deshacer_Add(ByVal Desc As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 16/10/06
'*************************************************
If frmMain.mnuUtilizarDeshacer.Checked = False Then Exit Sub

Dim i As Integer
Dim F As Integer
Dim j As Integer
' Desplazo todos los deshacer uno hacia atras
For i = maxDeshacer To 2 Step -1
    For F = XMinMapSize To XMaxMapSize
        For j = YMinMapSize To YMaxMapSize
            MapData_Deshacer(i, F, j) = MapData_Deshacer(i - 1, F, j)
        Next
    Next
    MapData_Deshacer_Info(i) = MapData_Deshacer_Info(i - 1)
Next
' Guardo los valores
For F = XMinMapSize To XMaxMapSize
    For j = YMinMapSize To YMaxMapSize
        MapData_Deshacer(1, F, j) = MapData(F, j)
    Next
Next
MapData_Deshacer_Info(1).Desc = Desc
MapData_Deshacer_Info(1).Libre = False
frmMain.mnuDeshacer.Caption = "&Deshacer (Ultimo: " & MapData_Deshacer_Info(1).Desc & ")"
frmMain.mnuDeshacer.Enabled = True
End Sub

''
' Deshacer un paso del Deshacer
'
Public Sub Deshacer_Recover()
'*************************************************
'Author: ^[GS]^
'Last modified: 15/10/06
'*************************************************
Dim i As Integer
Dim F As Integer
Dim j As Integer
Dim Body As Integer
Dim Head As Integer
Dim Heading As Byte
If MapData_Deshacer_Info(1).Libre = False Then
    ' Aplico deshacer
    For F = XMinMapSize To XMaxMapSize
        For j = YMinMapSize To YMaxMapSize
            If (MapData(F, j).NPCIndex <> 0 And MapData(F, j).NPCIndex <> MapData_Deshacer(1, F, j).NPCIndex) Or (MapData(F, j).NPCIndex <> 0 And MapData_Deshacer(1, F, j).NPCIndex = 0) Then
                ' Si ahi un NPC, y en el deshacer es otro lo borramos
                ' (o) Si aun no NPC y en el deshacer no esta
                MapData(F, j).NPCIndex = 0
                Call EraseChar(MapData(F, j).CharIndex)
            End If
            If MapData_Deshacer(1, F, j).NPCIndex <> 0 And MapData(F, j).NPCIndex = 0 Then
                ' Si ahi un NPC en el deshacer y en el no esta lo hacemos
                Body = NpcData(MapData_Deshacer(1, F, j).NPCIndex).Body
                Head = NpcData(MapData_Deshacer(1, F, j).NPCIndex).Head
                Heading = NpcData(MapData_Deshacer(1, F, j).NPCIndex).Heading
                Call MakeChar(NextOpenChar(), Body, Head, Heading, F, j)
            Else
                MapData(F, j) = MapData_Deshacer(1, F, j)
            End If
        Next
    Next
    MapData_Deshacer_Info(1).Libre = True
    ' Desplazo todos los deshacer uno hacia adelante
    For i = 1 To maxDeshacer - 1
        For F = XMinMapSize To XMaxMapSize
            For j = YMinMapSize To YMaxMapSize
                MapData_Deshacer(i, F, j) = MapData_Deshacer(i + 1, F, j)
            Next
        Next
        MapData_Deshacer_Info(i) = MapData_Deshacer_Info(i + 1)
    Next
    ' borro el ultimo
    MapData_Deshacer_Info(maxDeshacer).Libre = True
    ' ahi para deshacer?
    If MapData_Deshacer_Info(1).Libre = True Then
        frmMain.mnuDeshacer.Caption = "&Deshacer (no ahi nada que deshacer)"
        frmMain.mnuDeshacer.Enabled = False
    Else
        frmMain.mnuDeshacer.Caption = "&Deshacer (Ultimo: " & MapData_Deshacer_Info(1).Desc & ")"
        frmMain.mnuDeshacer.Enabled = True
    End If
Else
    MsgBox "No ahi acciones para deshacer", vbInformation
End If
End Sub

''
' Manda una advertencia de Edicion Critica
'
' @return   Nos devuelve si acepta o no el cambio

Private Function EditWarning() As Boolean
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If MsgBox(MSGDang, vbExclamation + vbYesNo) = vbNo Then
    EditWarning = True
Else
    EditWarning = False
End If
End Function


''
' Bloquea los Bordes del Mapa
'

Public Sub Bloquear_Bordes()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Dim Y As Integer
Dim X As Integer

If Not MapaCargado Then
    Exit Sub
End If

modEdicion.Deshacer_Add "Bloquear los bordes" ' Hago deshacer

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
        If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
            MapData(X, Y).Blocked = 1
        End If
    Next X
Next Y

'Set changed flag
MapInfo.Changed = 1
End Sub


''
' Coloca la superficie seleccionada al azar en el mapa
'

Public Sub Superficie_Azar()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

On Error Resume Next
Dim Y As Integer
Dim X As Integer
Dim Cuantos As Integer
Dim k As Integer

If Not MapaCargado Then
    Exit Sub
End If

Cuantos = InputBox("Cuantos Grh se deben poner en este mapa?", "Poner Grh Al Azar", 0)
If Cuantos > 0 Then
    modEdicion.Deshacer_Add "Insertar Superficie al Azar" ' Hago deshacer
    For k = 1 To Cuantos
        X = RandomNumber(10, 90)
        Y = RandomNumber(10, 90)
        If frmConfigSup.MOSAICO.value = vbChecked Then
          Dim aux As Integer
          Dim dy As Integer
          Dim dX As Integer
          If frmConfigSup.DespMosaic.value = vbChecked Then
                        dy = Val(frmConfigSup.DMLargo)
                        dX = Val(frmConfigSup.DMAncho.text)
          Else
                    dy = 0
                    dX = 0
          End If
                
          If frmMain.mnuAutoCompletarSuperficies.Checked = False Then
                aux = Val(frmMain.cGrh.text) + _
                (((Y + dy) Mod frmConfigSup.mLargo.text) * frmConfigSup.mAncho.text) + ((X + dX) Mod frmConfigSup.mAncho.text)
                If frmMain.cInsertarBloqueo.value = True Then
                    MapData(X, Y).Blocked = 1
                Else
                    MapData(X, Y).Blocked = 0
                End If
                MapData(X, Y).Graphic(Val(frmMain.cCapas.text)).GrhIndex = aux
                InitGrh MapData(X, Y).Graphic(Val(frmMain.cCapas.text)), aux
          Else
                Dim tXX As Integer, tYY As Integer, i As Integer, j As Integer, desptile As Integer
                tXX = X
                tYY = Y
                desptile = 0
                For i = 1 To frmConfigSup.mLargo.text
                    For j = 1 To frmConfigSup.mAncho.text
                        aux = Val(frmMain.cGrh.text) + desptile
                         
                        If frmMain.cInsertarBloqueo.value = True Then
                            MapData(tXX, tYY).Blocked = 1
                        Else
                            MapData(tXX, tYY).Blocked = 0
                        End If

                         MapData(tXX, tYY).Graphic(Val(frmMain.cCapas.text)).GrhIndex = aux
                         
                         InitGrh MapData(tXX, tYY).Graphic(Val(frmMain.cCapas.text)), aux
                         tXX = tXX + 1
                         desptile = desptile + 1
                    Next
                    tXX = X
                    tYY = tYY + 1
                Next
                tYY = Y
          End If
        End If
    Next
End If

'Set changed flag
MapInfo.Changed = 1

End Sub

''
' Coloca la superficie seleccionada en todos los bordes
'

Public Sub Superficie_Bordes()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

Dim Y As Integer
Dim X As Integer

If Not MapaCargado Then
    Exit Sub
End If

modEdicion.Deshacer_Add "Insertar Superficie en todos los bordes" ' Hago deshacer

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize

        If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then

          If frmConfigSup.MOSAICO.value = vbChecked Then
            Dim aux As Integer
            aux = Val(frmMain.cGrh.text) + _
            ((Y Mod frmConfigSup.mLargo) * frmConfigSup.mAncho) + (X Mod frmConfigSup.mAncho)
            If frmMain.cInsertarBloqueo.value = True Then
                MapData(X, Y).Blocked = 1
            Else
                MapData(X, Y).Blocked = 0
            End If
            MapData(X, Y).Graphic(Val(frmMain.cCapas.text)).GrhIndex = aux
            'Setup GRH
            InitGrh MapData(X, Y).Graphic(Val(frmMain.cCapas.text)), aux
          Else
            'Else Place graphic
            If frmMain.cInsertarBloqueo.value = True Then
                MapData(X, Y).Blocked = 1
            Else
                MapData(X, Y).Blocked = 0
            End If
            
            MapData(X, Y).Graphic(Val(frmMain.cCapas.text)).GrhIndex = Val(frmMain.cGrh.text)
            
            'Setup GRH
    
            InitGrh MapData(X, Y).Graphic(Val(frmMain.cCapas.text)), Val(frmMain.cGrh.text)
        End If
             'Erase NPCs
            If MapData(X, Y).NPCIndex > 0 Then
                EraseChar MapData(X, Y).CharIndex
                MapData(X, Y).NPCIndex = 0
            End If

            'Erase Objs
            MapData(X, Y).OBJInfo.OBJIndex = 0
            MapData(X, Y).OBJInfo.Amount = 0
            MapData(X, Y).ObjGrh.GrhIndex = 0

            'Clear exits
            MapData(X, Y).TileExit.Map = 0
            MapData(X, Y).TileExit.X = 0
            MapData(X, Y).TileExit.Y = 0

        End If

    Next X
Next Y

'Set changed flag
MapInfo.Changed = 1

End Sub

''
' Coloca la misma superficie seleccionada en todo el mapa
'

Public Sub Superficie_Todo()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

If EditWarning Then Exit Sub

Dim Y As Integer
Dim X As Integer

If Not MapaCargado Then
    Exit Sub
End If

modEdicion.Deshacer_Add "Insertar Superficie en todo el mapa" ' Hago deshacer

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize

        If frmConfigSup.MOSAICO.value = vbChecked Then
            Dim aux As Integer
            aux = Val(frmMain.cGrh.text) + _
            ((Y Mod frmConfigSup.mLargo) * frmConfigSup.mAncho) + (X Mod frmConfigSup.mAncho)
             MapData(X, Y).Graphic(Val(frmMain.cCapas.text)).GrhIndex = aux
            'Setup GRH
            InitGrh MapData(X, Y).Graphic(Val(frmMain.cCapas.text)), aux
        Else
            'Else Place graphic
            MapData(X, Y).Graphic(Val(frmMain.cCapas.text)).GrhIndex = Val(frmMain.cGrh.text)
            'Setup GRH
            InitGrh MapData(X, Y).Graphic(Val(frmMain.cCapas.text)), Val(frmMain.cGrh.text)
        End If

    Next X
Next Y

'Set changed flag
MapInfo.Changed = 1

End Sub
Public Sub Luces_Todo()
'*************************************************
'Author: Lorwik
'Last modified: 19/11/2011
'*************************************************

If EditWarning Then Exit Sub

Dim Y As Integer
Dim X As Integer

If Not MapaCargado Then
    Exit Sub
End If

modEdicion.Deshacer_Add "Insertar Luces en todo el mapa" ' Hago deshacer

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
    Light.Create_Light_To_Map X, Y, frmMain.cRango, Val(frmMain.R), Val(frmMain.G), Val(frmMain.B)
    Next X
Next Y

'Set changed flag
MapInfo.Changed = 1

End Sub

''
' Modifica los bloqueos de todo mapa
'
' @param Valor Especifica el estado de Bloqueo que se asignara


Public Sub Bloqueo_Todo(ByVal Valor As Byte)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

If EditWarning Then Exit Sub


Dim Y As Integer
Dim X As Integer

If Not MapaCargado Then
    Exit Sub
End If

modEdicion.Deshacer_Add "Bloquear todo el mapa" ' Hago deshacer

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
        MapData(X, Y).Blocked = Valor
    Next X
Next Y

'Set changed flag
MapInfo.Changed = 1

End Sub

''
' Borra todo el Mapa menos los Triggers
'

Public Sub Borrar_Mapa()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

If EditWarning Then Exit Sub


Dim Y As Integer
Dim X As Integer

If Not MapaCargado Then
    Exit Sub
End If

modEdicion.Deshacer_Add "Borrar todo el mapa menos Triggers" ' Hago deshacer

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
        MapData(X, Y).Graphic(1).GrhIndex = 1
        'Change blockes status
        MapData(X, Y).Blocked = 0

        'Erase layer 2 and 3
        MapData(X, Y).Graphic(2).GrhIndex = 0
        MapData(X, Y).Graphic(3).GrhIndex = 0
        MapData(X, Y).Graphic(4).GrhIndex = 0

        'Erase NPCs
        If MapData(X, Y).NPCIndex > 0 Then
            EraseChar MapData(X, Y).CharIndex
            MapData(X, Y).NPCIndex = 0
        End If

        'Erase Objs
        MapData(X, Y).OBJInfo.OBJIndex = 0
        MapData(X, Y).OBJInfo.Amount = 0
        MapData(X, Y).ObjGrh.GrhIndex = 0

        'Clear exits
        MapData(X, Y).TileExit.Map = 0
        MapData(X, Y).TileExit.X = 0
        MapData(X, Y).TileExit.Y = 0
        
        InitGrh MapData(X, Y).Graphic(1), 1

    Next X
Next Y

'Set changed flag
MapInfo.Changed = 1
End Sub

''
' Elimita los NPCs del mapa
'
' @param Hostiles Indica si elimita solo hostiles o solo npcs no hostiles

Public Sub Quitar_NPCs(ByVal Hostiles As Boolean)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

If EditWarning Then Exit Sub

modEdicion.Deshacer_Add "Quitar todos los NPCs" & IIf(Hostiles = True, " Hostiles", "") ' Hago deshacer

Dim Y As Integer
Dim X As Integer

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
        If MapData(X, Y).NPCIndex > 0 Then
            If (Hostiles = True And MapData(X, Y).NPCIndex >= 500) Or (Hostiles = False And MapData(X, Y).NPCIndex < 500) Then
                Call EraseChar(MapData(X, Y).CharIndex)
                MapData(X, Y).NPCIndex = 0
            End If
        End If
    Next X
Next Y

Call DibujarMiniMapa ' Radar

'Set changed flag
MapInfo.Changed = 1
End Sub

''
' Elimita todos los Objetos del mapa
'

Public Sub Quitar_Objetos()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

If EditWarning Then Exit Sub

modEdicion.Deshacer_Add "Quitar todos los Objetos" ' Hago deshacer

Dim Y As Integer
Dim X As Integer

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
        If MapData(X, Y).OBJInfo.OBJIndex > 0 Then
            If MapData(X, Y).Graphic(3).GrhIndex = MapData(X, Y).ObjGrh.GrhIndex Then MapData(X, Y).Graphic(3).GrhIndex = 0
            MapData(X, Y).OBJInfo.OBJIndex = 0
            MapData(X, Y).OBJInfo.Amount = 0
        End If
    Next X
Next Y

'Set changed flag
MapInfo.Changed = 1
End Sub

''
' Elimina todos los Triggers del mapa
'

Public Sub Quitar_Triggers()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

If EditWarning Then Exit Sub

modEdicion.Deshacer_Add "Quitar todos los Triggers" ' Hago deshacer

Dim Y As Integer
Dim X As Integer

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
        If MapData(X, Y).Trigger > 0 Then
            MapData(X, Y).Trigger = 0
        End If
    Next X
Next Y

'Set changed flag
MapInfo.Changed = 1
End Sub

''
' Elimita todos los translados del mapa
'

Public Sub Quitar_Translados()
'*************************************************
'Author: ^[GS]^
'Last modified: 16/10/06
'*************************************************

If EditWarning Then Exit Sub

modEdicion.Deshacer_Add "Quitar todos los Translados" ' Hago deshacer

Dim Y As Integer
Dim X As Integer

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
        If MapData(X, Y).TileExit.Map > 0 Then
            MapData(X, Y).TileExit.Map = 0
            MapData(X, Y).TileExit.X = 0
            MapData(X, Y).TileExit.Y = 0
        End If
    Next X
Next Y

'Set changed flag
MapInfo.Changed = 1

End Sub

Public Sub Quitar_Luces()
'*************************************************
'Author: Lorwik
'Last modified: 19/11/2011
'*************************************************

If EditWarning Then Exit Sub

modEdicion.Deshacer_Add "Quitar todas las luces" ' Hago deshacer

Dim Y As Integer
Dim X As Integer
Dim i As Byte

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
            If MapData(X, Y).base_light(0) Or MapData(X, Y).base_light(1) _
                Or MapData(X, Y).base_light(2) Or MapData(X, Y).base_light(3) Then
                MapData(X, Y).light_value(0) = 0
                MapData(X, Y).light_value(1) = 0
                MapData(X, Y).light_value(2) = 0
                MapData(X, Y).light_value(3) = 0
                MapData(X, Y).base_light(0) = 0
                MapData(X, Y).base_light(1) = 0
                MapData(X, Y).base_light(2) = 0
                MapData(X, Y).base_light(3) = 0
                engine.Particle_Group_Remove_All
            End If
    Next X
Next Y

'Set changed flag
MapInfo.Changed = 1

End Sub

''
' Elimita todo lo que se encuentre en los bordes del mapa
'

Public Sub Quitar_Bordes()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

If EditWarning Then Exit Sub

'*****************************************************************
'Clears a border in a room with current GRH
'*****************************************************************

Dim Y As Integer
Dim X As Integer

If Not MapaCargado Then
    Exit Sub
End If

modEdicion.Deshacer_Add "Quitar todos los Bordes" ' Hago deshacer

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize

        If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
        
            MapData(X, Y).Graphic(1).GrhIndex = 1
            InitGrh MapData(X, Y).Graphic(1), 1
            MapData(X, Y).Blocked = 0
            
             'Erase NPCs
            If MapData(X, Y).NPCIndex > 0 Then
                EraseChar MapData(X, Y).CharIndex
                MapData(X, Y).NPCIndex = 0
            End If

            'Erase Objs
            MapData(X, Y).OBJInfo.OBJIndex = 0
            MapData(X, Y).OBJInfo.Amount = 0
            MapData(X, Y).ObjGrh.GrhIndex = 0

            'Clear exits
            MapData(X, Y).TileExit.Map = 0
            MapData(X, Y).TileExit.X = 0
            MapData(X, Y).TileExit.Y = 0
            
            ' Triggers
            MapData(X, Y).Trigger = 0

        End If

    Next X
Next Y

'Set changed flag
MapInfo.Changed = 1

End Sub

''
' Elimita una capa completa del mapa
'
' @param Capa Especifica la capa


Public Sub Quitar_Capa(ByVal Capa As Byte)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

If EditWarning Then Exit Sub

'*****************************************************************
'Clears one layer
'*****************************************************************

Dim Y As Integer
Dim X As Integer

If Not MapaCargado Then
    Exit Sub
End If
modEdicion.Deshacer_Add "Quitar Capa " & Capa ' Hago deshacer

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
        If Capa = 1 Then
            MapData(X, Y).Graphic(Capa).GrhIndex = 1
        Else
            MapData(X, Y).Graphic(Capa).GrhIndex = 0
        End If
    Next X
Next Y

'Set changed flag
MapInfo.Changed = 1
End Sub

''
' Acciona la operacion al hacer doble click en una posicion del mapa
'
' @param tX Especifica la posicion X en el mapa
' @param tY Espeficica la posicion Y en el mapa

Sub DobleClick(tx As Byte, ty As Byte)
'*************************************************
'Author: ^[GS]^
'Last modified: 01/11/08
'*************************************************
' Selecciones
Seleccionando = False ' GS
SeleccionIX = 0
SeleccionIY = 0
SeleccionFX = 0
SeleccionFY = 0
' Translados
Dim tTrans As WorldPos
tTrans = MapData(tx, ty).TileExit
If tTrans.Map > 0 Then
    If LenB(frmMain.Dialog.FileName) <> 0 Then
        If FileExist(PATH_Save & NameMap_Save & tTrans.Map & ".map", vbArchive) = True Then
            Call modMapIO.NuevoMapa
            frmMain.Dialog.FileName = PATH_Save & NameMap_Save & tTrans.Map & ".map"
            modMapIO.AbrirMapa frmMain.Dialog.FileName
            UserPos.X = tTrans.X
            UserPos.Y = tTrans.Y
            If WalkMode = True Then
                MoveCharbyPos UserCharIndex, UserPos.X, UserPos.Y
                charlist(UserCharIndex).Heading = SOUTH
            End If
            frmMain.mnuReAbrirMapa.Enabled = True
        End If
    End If
End If
End Sub

''
' Realiza una operacion de edicion aislada sobre el mapa
'
' @param Button Indica el estado del Click del mouse
' @param tX Especifica la posicion X en el mapa
' @param tY Especifica la posicion Y en el mapa

Sub ClickEdit(Button As Integer, tx As Byte, ty As Byte)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

    Dim loopc As Integer
    Dim NPCIndex As Integer
    Dim OBJIndex As Integer
    Dim Head As Integer
    Dim Body As Integer
    Dim Heading As Byte
    
    If ty < 1 Or ty > 100 Then Exit Sub
    If tx < 1 Or tx > 100 Then Exit Sub
    
    
    If Button = 0 Then
        ' Pasando sobre :P
        SobreY = ty
        SobreX = tx
        
    End If
    
    'Right
    
    If Button = vbRightButton Then
        ' Posicion
        frmMain.StatTxt.text = frmMain.StatTxt.text & ENDL & ENDL & "Posición " & tx & "," & ty
        
        ' Bloqueos
        If MapData(tx, ty).Blocked = 1 Then frmMain.StatTxt.text = frmMain.StatTxt.text & " (BLOQ)"
        
        ' Translados
        If MapData(tx, ty).TileExit.Map > 0 Then
            If frmMain.mnuAutoCapturarTranslados.Checked = True Then
                frmMain.tTMapa.text = MapData(tx, ty).TileExit.Map
                frmMain.tTX.text = MapData(tx, ty).TileExit.X
                frmMain.tTY = MapData(tx, ty).TileExit.Y
            End If
            frmMain.StatTxt.text = frmMain.StatTxt.text & " (Trans.: " & MapData(tx, ty).TileExit.Map & "," & MapData(tx, ty).TileExit.X & "," & MapData(tx, ty).TileExit.Y & ")"
        End If
        
        ' NPCs
        If MapData(tx, ty).NPCIndex > 0 Then
            If MapData(tx, ty).NPCIndex > 499 Then
                frmMain.StatTxt.text = frmMain.StatTxt.text & " (NPC-Hostil: " & MapData(tx, ty).NPCIndex & " - " & NpcData(MapData(tx, ty).NPCIndex).name & ")"
            Else
                frmMain.StatTxt.text = frmMain.StatTxt.text & " (NPC: " & MapData(tx, ty).NPCIndex & " - " & NpcData(MapData(tx, ty).NPCIndex).name & ")"
            End If
        End If
        
        ' OBJs
        If MapData(tx, ty).OBJInfo.OBJIndex > 0 Then
            frmMain.StatTxt.text = frmMain.StatTxt.text & " (Obj: " & MapData(tx, ty).OBJInfo.OBJIndex & " - " & ObjData(MapData(tx, ty).OBJInfo.OBJIndex).name & " - Cant.:" & MapData(tx, ty).OBJInfo.Amount & ")"
        End If
        
        ' Capas
        frmMain.StatTxt.text = frmMain.StatTxt.text & ENDL & "Capa1: " & MapData(tx, ty).Graphic(1).GrhIndex & " - Capa2: " & MapData(tx, ty).Graphic(2).GrhIndex & " - Capa3: " & MapData(tx, ty).Graphic(3).GrhIndex & " - Capa4: " & MapData(tx, ty).Graphic(4).GrhIndex
        If frmMain.mnuAutoCapturarSuperficie.Checked = True And frmMain.cSeleccionarSuperficie.value = False Then
            If MapData(tx, ty).Graphic(4).GrhIndex <> 0 Then
                frmMain.cCapas.text = 4
                frmMain.cGrh.text = MapData(tx, ty).Graphic(4).GrhIndex
            ElseIf MapData(tx, ty).Graphic(3).GrhIndex <> 0 Then
                frmMain.cCapas.text = 3
                frmMain.cGrh.text = MapData(tx, ty).Graphic(3).GrhIndex
            ElseIf MapData(tx, ty).Graphic(2).GrhIndex <> 0 Then
                frmMain.cCapas.text = 2
                frmMain.cGrh.text = MapData(tx, ty).Graphic(2).GrhIndex
            ElseIf MapData(tx, ty).Graphic(1).GrhIndex <> 0 Then
                frmMain.cCapas.text = 1
                frmMain.cGrh.text = MapData(tx, ty).Graphic(1).GrhIndex
            End If
        End If
        
        ' Limpieza
        If Len(frmMain.StatTxt.text) > 4000 Then
            frmMain.StatTxt.text = Right(frmMain.StatTxt.text, 3000)
        End If
        frmMain.StatTxt.SelStart = Len(frmMain.StatTxt.text)
        
        Exit Sub
    End If
    
    
    'Left click
    If Button = vbLeftButton Then
            
            'Erase 2-3
            If frmMain.cQuitarEnTodasLasCapas.value = True Then
                modEdicion.Deshacer_Add "Quitar Todas las Capas (2/3)" ' Hago deshacer
                MapInfo.Changed = 1 'Set changed flag
                For loopc = 2 To 3
                    MapData(tx, ty).Graphic(loopc).GrhIndex = 0
                Next loopc
                
                Exit Sub
            End If
    
            'Borrar "esta" Capa
            If frmMain.cQuitarEnEstaCapa.value = True Then
                If Val(frmMain.cCapas.text) = 1 Then
                    If MapData(tx, ty).Graphic(1).GrhIndex <> 1 Then
                        modEdicion.Deshacer_Add "Quitar Capa 1" ' Hago deshacer
                        MapInfo.Changed = 1 'Set changed flag
                        MapData(tx, ty).Graphic(1).GrhIndex = 1
                        Exit Sub
                    End If
                ElseIf MapData(tx, ty).Graphic(Val(frmMain.cCapas.text)).GrhIndex <> 0 Then
                    modEdicion.Deshacer_Add "Quitar Capa " & frmMain.cCapas.text  ' Hago deshacer
                    MapInfo.Changed = 1 'Set changed flag
                    MapData(tx, ty).Graphic(Val(frmMain.cCapas.text)).GrhIndex = 0
                    Exit Sub
                End If
            End If
    
        '************** Place grh
        If frmMain.cSeleccionarSuperficie.value = True Then
            
            If frmConfigSup.MOSAICO.value = vbChecked Then
              Dim aux As Integer
              Dim dy As Integer
              Dim dX As Integer
              If frmConfigSup.DespMosaic.value = vbChecked Then
                            dy = Val(frmConfigSup.DMLargo)
                            dX = Val(frmConfigSup.DMAncho.text)
              Else
                        dy = 0
                        dX = 0
              End If
                    
              If frmMain.mnuAutoCompletarSuperficies.Checked = False Then
                    modEdicion.Deshacer_Add "Insertar Superficie' Hago deshacer"
                    MapInfo.Changed = 1 'Set changed flag
                    aux = Val(frmMain.cGrh.text) + _
                    (((ty + dy) Mod frmConfigSup.mLargo.text) * frmConfigSup.mAncho.text) + ((tx + dX) Mod frmConfigSup.mAncho.text)
                     If MapData(tx, ty).Graphic(Val(frmMain.cCapas.text)).GrhIndex <> aux Or MapData(tx, ty).Blocked <> frmMain.SelectPanel(2).value Then
                        MapData(tx, ty).Graphic(Val(frmMain.cCapas.text)).GrhIndex = aux
                        InitGrh MapData(tx, ty).Graphic(Val(frmMain.cCapas.text)), aux
                    End If
              Else
                modEdicion.Deshacer_Add "Insertar Auto-Completar Superficie' Hago deshacer"
                MapInfo.Changed = 1 'Set changed flag
                Dim tXX As Integer, tYY As Integer, i As Integer, j As Integer, desptile As Integer
                tXX = tx
                tYY = ty
                desptile = 0
                For i = 1 To frmConfigSup.mLargo.text
                    For j = 1 To frmConfigSup.mAncho.text
                        aux = Val(frmMain.cGrh.text) + desptile
                        MapData(tXX, tYY).Graphic(Val(frmMain.cCapas.text)).GrhIndex = aux
                        InitGrh MapData(tXX, tYY).Graphic(Val(frmMain.cCapas.text)), aux
                        tXX = tXX + 1
                        desptile = desptile + 1
                    Next
                    tXX = tx
                    tYY = tYY + 1
                Next
                tYY = ty
                    
                    
              End If
              
            Else
                'Else Place graphic
                If MapData(tx, ty).Blocked <> frmMain.SelectPanel(2).value Or MapData(tx, ty).Graphic(Val(frmMain.cCapas.text)).GrhIndex <> Val(frmMain.cGrh.text) Then
                    modEdicion.Deshacer_Add "Quitar Superficie en Capa " & frmMain.cCapas.text ' Hago deshacer
                    MapInfo.Changed = 1 'Set changed flag
                    MapData(tx, ty).Graphic(Val(frmMain.cCapas.text)).GrhIndex = Val(frmMain.cGrh.text)
                    'Setup GRH
                    InitGrh MapData(tx, ty).Graphic(Val(frmMain.cCapas.text)), Val(frmMain.cGrh.text)
                End If
            End If
            
        End If
        '************** Place blocked tile
        If frmMain.cInsertarBloqueo.value = True Then
            If MapData(tx, ty).Blocked <> 1 Then
                modEdicion.Deshacer_Add "Insertar Bloqueo" ' Hago deshacer
                MapInfo.Changed = 1 'Set changed flag
                MapData(tx, ty).Blocked = 1
            End If
        ElseIf frmMain.cQuitarBloqueo.value = True Then
            If MapData(tx, ty).Blocked <> 0 Then
                modEdicion.Deshacer_Add "Quitar Bloqueo" ' Hago deshacer
                MapInfo.Changed = 1 'Set changed flag
                MapData(tx, ty).Blocked = 0
            End If
        End If
    
        '************** Place exit
        If frmMain.cInsertarTrans.value = True Then
            If Cfg_TrOBJ > 0 And Cfg_TrOBJ <= NumOBJs And frmMain.cInsertarTransOBJ.value = True Then
                If ObjData(Cfg_TrOBJ).ObjType = 19 Then
                    modEdicion.Deshacer_Add "Insertar Objeto de Translado" ' Hago deshacer
                    MapInfo.Changed = 1 'Set changed flag
                    InitGrh MapData(tx, ty).ObjGrh, ObjData(Cfg_TrOBJ).GrhIndex
                    MapData(tx, ty).OBJInfo.OBJIndex = Cfg_TrOBJ
                    MapData(tx, ty).OBJInfo.Amount = 1
                End If
            End If
            If Val(frmMain.tTMapa.text) < 0 Or Val(frmMain.tTMapa.text) > 9000 Then
                MsgBox "Valor de Mapa invalido", vbCritical + vbOKOnly
                Exit Sub
            ElseIf Val(frmMain.tTX.text) < 0 Or Val(frmMain.tTX.text) > 100 Then
                MsgBox "Valor de X invalido", vbCritical + vbOKOnly
                Exit Sub
            ElseIf Val(frmMain.tTY.text) < 0 Or Val(frmMain.tTY.text) > 100 Then
                MsgBox "Valor de Y invalido", vbCritical + vbOKOnly
                Exit Sub
            End If
                If frmMain.cUnionManual.value = True Then
                    modEdicion.Deshacer_Add "Insertar Translado de Union Manual' Hago deshacer"
                    MapInfo.Changed = 1 'Set changed flag
                    MapData(tx, ty).TileExit.Map = Val(frmMain.tTMapa.text)
                    If tx >= 90 Then ' 21 ' derecha
                              MapData(tx, ty).TileExit.X = 12
                              MapData(tx, ty).TileExit.Y = ty
                    ElseIf tx <= 11 Then ' 9 ' izquierda
                        MapData(tx, ty).TileExit.X = 91
                        MapData(tx, ty).TileExit.Y = ty
                    End If
                    If ty >= 91 Then ' 94 '''' hacia abajo
                             MapData(tx, ty).TileExit.Y = 11
                             MapData(tx, ty).TileExit.X = tx
                    ElseIf ty <= 10 Then ''' hacia arriba
                        MapData(tx, ty).TileExit.Y = 90
                        MapData(tx, ty).TileExit.X = tx
                    End If
                Else
                    modEdicion.Deshacer_Add "Insertar Translado" ' Hago deshacer
                    MapInfo.Changed = 1 'Set changed flag
                    MapData(tx, ty).TileExit.Map = Val(frmMain.tTMapa.text)
                    MapData(tx, ty).TileExit.X = Val(frmMain.tTX.text)
                    MapData(tx, ty).TileExit.Y = Val(frmMain.tTY.text)
                End If
        ElseIf frmMain.cQuitarTrans.value = True Then
                modEdicion.Deshacer_Add "Quitar Translado" ' Hago deshacer
                MapInfo.Changed = 1 'Set changed flag
                MapData(tx, ty).TileExit.Map = 0
                MapData(tx, ty).TileExit.X = 0
                MapData(tx, ty).TileExit.Y = 0
        End If
    
        '************** Place NPC
        If frmMain.cInsertarFunc(0).value = True Then
            If frmMain.cNumFunc(0).text > 0 Then
                NPCIndex = frmMain.cNumFunc(0).text
                If NPCIndex <> MapData(tx, ty).NPCIndex Then
                    modEdicion.Deshacer_Add "Insertar NPC" ' Hago deshacer
                    MapInfo.Changed = 1 'Set changed flag
                    Body = NpcData(NPCIndex).Body
                    Head = NpcData(NPCIndex).Head
                    Heading = NpcData(NPCIndex).Heading
                    Call MakeChar(NextOpenChar(), Body, Head, Heading, tx, ty)
                    MapData(tx, ty).NPCIndex = NPCIndex
                End If
            End If
        ElseIf frmMain.cInsertarFunc(1).value = True Then
            If frmMain.cNumFunc(1).text > 0 Then
                NPCIndex = frmMain.cNumFunc(1).text
                If NPCIndex <> (MapData(tx, ty).NPCIndex) Then
                    modEdicion.Deshacer_Add "Insertar NPC Hostil' Hago deshacer"
                    MapInfo.Changed = 1 'Set changed flag
                    Body = NpcData(NPCIndex).Body
                    Head = NpcData(NPCIndex).Head
                    Heading = NpcData(NPCIndex).Heading
                    Call MakeChar(NextOpenChar(), Body, Head, Heading, tx, ty)
                    MapData(tx, ty).NPCIndex = NPCIndex
                End If
            End If
        ElseIf frmMain.cQuitarFunc(0).value = True Or frmMain.cQuitarFunc(1).value = True Then
            If MapData(tx, ty).NPCIndex > 0 Then
                modEdicion.Deshacer_Add "Quitar NPC" ' Hago deshacer
                MapInfo.Changed = 1 'Set changed flag
                MapData(tx, ty).NPCIndex = 0
                Call EraseChar(MapData(tx, ty).CharIndex)
            End If
        End If
    
        ' ***************** Control de Funcion de Objetos *****************
        If frmMain.cInsertarFunc(2).value = True Then ' Insertar Objeto
            If frmMain.cNumFunc(2).text > 0 Then
                OBJIndex = frmMain.cNumFunc(2).text
                If MapData(tx, ty).OBJInfo.OBJIndex <> OBJIndex Or MapData(tx, ty).OBJInfo.Amount <> Val(frmMain.cCantFunc(2).text) Then
                    modEdicion.Deshacer_Add "Insertar Objeto" ' Hago deshacer
                    MapInfo.Changed = 1 'Set changed flag
                    InitGrh MapData(tx, ty).ObjGrh, ObjData(OBJIndex).GrhIndex
                    MapData(tx, ty).OBJInfo.OBJIndex = OBJIndex
                    MapData(tx, ty).OBJInfo.Amount = Val(frmMain.cCantFunc(2).text)
                    Select Case ObjData(OBJIndex).ObjType
                        Case 4, 8, 10, 22 ' Arboles, Carteles, Foros, Yacimientos
                            MapData(tx, ty).Graphic(3) = MapData(tx, ty).ObjGrh
                    End Select
                End If
            End If
        ElseIf frmMain.cQuitarFunc(2).value = True Then ' Quitar Objeto
            If MapData(tx, ty).OBJInfo.OBJIndex <> 0 Or MapData(tx, ty).OBJInfo.Amount <> 0 Then
                modEdicion.Deshacer_Add "Quitar Objeto" ' Hago deshacer
                MapInfo.Changed = 1 'Set changed flag
                If MapData(tx, ty).Graphic(3).GrhIndex = MapData(tx, ty).ObjGrh.GrhIndex Then MapData(tx, ty).Graphic(3).GrhIndex = 0
                MapData(tx, ty).ObjGrh.GrhIndex = 0
                MapData(tx, ty).OBJInfo.OBJIndex = 0
                MapData(tx, ty).OBJInfo.Amount = 0
            End If
        End If
        
        ' ***************** Control de Funcion de Triggers *****************
        If frmMain.cInsertarTrigger.value = True Then ' Insertar Trigger
            If MapData(tx, ty).Trigger <> frmMain.lListado(4).ListIndex Then
                modEdicion.Deshacer_Add "Insertar Trigger" ' Hago deshacer
                MapInfo.Changed = 1 'Set changed flag
                MapData(tx, ty).Trigger = frmMain.lListado(4).ListIndex
            End If
        ElseIf frmMain.cQuitarTrigger.value = True Then ' Quitar Trigger
            If MapData(tx, ty).Trigger <> 0 Then
                modEdicion.Deshacer_Add "Quitar Trigger" ' Hago deshacer
                MapInfo.Changed = 1 'Set changed flag
                MapData(tx, ty).Trigger = 0
            End If
        End If
        
        ' ***************** Control de Funcion de Particles! *****************
        If frmMain.cmdAdd.value = True Then ' Insertar Particle
            modEdicion.Deshacer_Add "Insertar Particle"
            MapInfo.Changed = 1 'Set changed flag
            General_Particle_Create frmMain.lstParticle.ListIndex + 1, tx, ty, frmMain.Life.text
            MapData(tx, ty).Particle_Index = frmMain.lstParticle.ListIndex + 1
        ElseIf frmMain.cmdDel.value = True Then ' Quitar Particle
            If MapData(tx, ty).particle_group_index <> 0 Then
                modEdicion.Deshacer_Add "Quitar Particle" ' Hago deshacer
                MapInfo.Changed = 1 'Set changed flag
                engine.Particle_Group_Remove MapData(tx, ty).particle_group_index
                'MapData(tx, ty).particle_group_index = 0
                'MapData(tx, ty).Particle_Index = 0
            End If
        End If
        
               '*****************LUCES******************************
        If frmMain.cInsertarLuz.value Then
            If Val(frmMain.cRango = 0) Then Exit Sub
            Light.Create_Light_To_Map tx, ty, frmMain.cRango, Val(frmMain.R), Val(frmMain.G), Val(frmMain.B)
            MapInfo.Changed = 1 'Set changed flag
        ElseIf frmMain.cQuitarLuz.value Then
            Light.Delete_Light_To_Map tx, ty
            MapInfo.Changed = 1 'Set changed flag
        End If
    End If

End Sub
