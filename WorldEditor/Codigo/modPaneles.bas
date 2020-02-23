Attribute VB_Name = "modPaneles"
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
' modPaneles
'
' @remarks Funciones referentes a los Paneles de Funcion
' @author gshaxor@gmail.com
' @version 0.3.28
' @date 20060530

Option Explicit

''
' Activa/Desactiva el Estado de la Funcion en el Panel Superior
'
' @param Numero Especifica en numero de funcion
' @param Activado Especifica si esta o no activado

Public Sub EstSelectPanel(ByVal Numero As Byte, ByVal Activado As Boolean)
'*************************************************
'Author: ^[GS]^
'Last modified: 30/05/06
'*************************************************
    If Activado = True Then
        frmMain.SelectPanel(Numero).GradientMode = lv_Bottom2Top
        frmMain.SelectPanel(Numero).HoverBackColor = frmMain.SelectPanel(Numero).GradientColor
        If frmMain.mnuVerAutomatico.Checked = True Then
            Select Case Numero
                Case 0
                    If frmMain.cCapas.text = 4 Then
                        frmMain.mnuVerCapa4.Tag = CInt(frmMain.mnuVerCapa4.Checked)
                        frmMain.mnuVerCapa4.Checked = True
                    ElseIf frmMain.cCapas.text = 3 Then
                        frmMain.mnuVerCapa3.Tag = CInt(frmMain.mnuVerCapa3.Checked)
                        frmMain.mnuVerCapa3.Checked = True
                    ElseIf frmMain.cCapas.text = 2 Then
                        frmMain.mnuVerCapa2.Tag = CInt(frmMain.mnuVerCapa2.Checked)
                        frmMain.mnuVerCapa2.Checked = True
                    End If
                Case 2
                    frmMain.cVerBloqueos.Tag = CInt(frmMain.cVerBloqueos.value)
                    frmMain.cVerBloqueos.value = True
                    frmMain.mnuVerBloqueos.Checked = frmMain.cVerBloqueos.value
                Case 6
                    frmMain.cVerTriggers.Tag = CInt(frmMain.cVerTriggers.value)
                    frmMain.cVerTriggers.value = True
                    frmMain.mnuVerTriggers.Checked = frmMain.cVerTriggers.value
            End Select
        End If
    Else
        frmMain.SelectPanel(Numero).HoverBackColor = frmMain.SelectPanel(Numero).BackColor
        frmMain.SelectPanel(Numero).GradientMode = lv_NoGradient
        If frmMain.mnuVerAutomatico.Checked = True Then
            Select Case Numero
                Case 0
                    If frmMain.cCapas.text = 4 Then
                        If LenB(frmMain.mnuVerCapa3.Tag) <> 0 Then frmMain.mnuVerCapa4.Checked = CBool(frmMain.mnuVerCapa4.Tag)
                    ElseIf frmMain.cCapas.text = 3 Then
                        If LenB(frmMain.mnuVerCapa3.Tag) <> 0 Then frmMain.mnuVerCapa3.Checked = CBool(frmMain.mnuVerCapa3.Tag)
                    ElseIf frmMain.cCapas.text = 2 Then
                        If LenB(frmMain.mnuVerCapa2.Tag) <> 0 Then frmMain.mnuVerCapa2.Checked = CBool(frmMain.mnuVerCapa2.Tag)
                    End If
                Case 2
                    If LenB(frmMain.cVerBloqueos.Tag) = 0 Then frmMain.cVerBloqueos.Tag = 0
                    frmMain.cVerBloqueos.value = CBool(frmMain.cVerBloqueos.Tag)
                    frmMain.mnuVerBloqueos.Checked = frmMain.cVerBloqueos.value
                Case 6
                    If LenB(frmMain.cVerTriggers.Tag) = 0 Then frmMain.cVerTriggers.Tag = 0
                    frmMain.cVerTriggers.value = CBool(frmMain.cVerTriggers.Tag)
                    frmMain.mnuVerTriggers.Checked = frmMain.cVerTriggers.value
            End Select
        End If
    End If
End Sub

''
' Muestra los controles que componen a la funcion seleccionada del Panel
'
' @param Numero Especifica el numero de Funcion
' @param Ver Especifica si se va a ver o no
' @param Normal Inidica que ahi que volver todo No visible

Public Sub VerFuncion(ByVal Numero As Byte, ByVal Ver As Boolean, Optional Normal As Boolean)
'*************************************************
'Author: ^[GS]^
'Last modified: 24/11/08
'*************************************************
    If Normal = True Then
        Call VerFuncion(vMostrando, False, False)
    End If
    Select Case Numero
        Case 0 ' Superficies
            frmMain.lListado(0).Visible = Ver
            frmMain.cFiltro(0).Visible = Ver
            frmMain.cCapas.Visible = Ver
            frmMain.cGrh.Visible = Ver
            frmMain.cQuitarEnEstaCapa.Visible = Ver
            frmMain.cQuitarEnTodasLasCapas.Visible = Ver
            frmMain.cSeleccionarSuperficie.Visible = Ver
            frmMain.lbFiltrar(0).Visible = Ver
            frmMain.lbCapas.Visible = Ver
            frmMain.lbGrh.Visible = Ver
            frmMain.PreviewGrh.Visible = Ver
            If Ver = True Then
                frmMain.StatTxt.Top = 672
                frmMain.StatTxt.Height = 37
            Else
                frmMain.StatTxt.Top = 416
                frmMain.StatTxt.Height = 293
            End If
        Case 1 ' Translados
            frmMain.lMapN.Visible = Ver
            frmMain.lXhor.Visible = Ver
            frmMain.lYver.Visible = Ver
            frmMain.tTMapa.Visible = Ver
            frmMain.tTX.Visible = Ver
            frmMain.tTY.Visible = Ver
            frmMain.cInsertarTrans.Visible = Ver
            frmMain.cInsertarTransOBJ.Visible = Ver
            frmMain.cUnionManual.Visible = Ver
            frmMain.cUnionAuto.Visible = Ver
            frmMain.cQuitarTrans.Visible = Ver
        Case 2 ' Bloqueos
            frmMain.cQuitarBloqueo.Visible = Ver
            frmMain.cInsertarBloqueo.Visible = Ver
            frmMain.cVerBloqueos.Visible = Ver
        Case 3  ' NPCs
            frmMain.lListado(1).Visible = Ver
            frmMain.cFiltro(1).Visible = Ver
            frmMain.lbFiltrar(1).Visible = Ver
            frmMain.lNumFunc(Numero - 3).Visible = Ver
            frmMain.cNumFunc(Numero - 3).Visible = Ver
            frmMain.cInsertarFunc(Numero - 3).Visible = Ver
            frmMain.cQuitarFunc(Numero - 3).Visible = Ver
            frmMain.cAgregarFuncalAzar(Numero - 3).Visible = Ver
            frmMain.lCantFunc(Numero - 3).Visible = Ver
            frmMain.cCantFunc(Numero - 3).Visible = Ver
        Case 4 ' NPCs Hostiles
            'frmMain.lListado(1).Visible = Ver
            'frmMain.cFiltro(1).Visible = Ver
            'frmMain.lbFiltrar(1).Visible = Ver
            'frmMain.lNumFunc(Numero - 3).Visible = Ver
            'frmMain.cNumFunc(Numero - 3).Visible = Ver
            'frmMain.cInsertarFunc(Numero - 3).Visible = Ver
            'frmMain.cQuitarFunc(Numero - 3).Visible = Ver
            'frmMain.cAgregarFuncalAzar(Numero - 3).Visible = Ver
            'frmMain.lCantFunc(Numero - 3).Visible = Ver
            'frmMain.cCantFunc(Numero - 3).Visible = Ver
        Case 5 ' OBJs
            frmMain.lListado(3).Visible = Ver
            frmMain.cFiltro(3).Visible = Ver
            frmMain.lbFiltrar(3).Visible = Ver
            frmMain.lNumFunc(Numero - 3).Visible = Ver
            frmMain.cNumFunc(Numero - 3).Visible = Ver
            frmMain.cInsertarFunc(Numero - 3).Visible = Ver
            frmMain.cQuitarFunc(Numero - 3).Visible = Ver
            frmMain.cAgregarFuncalAzar(Numero - 3).Visible = Ver
            frmMain.lCantFunc(Numero - 3).Visible = Ver
            frmMain.cCantFunc(Numero - 3).Visible = Ver
        Case 6 ' Triggers
            frmMain.cQuitarTrigger.Visible = Ver
            frmMain.cInsertarTrigger.Visible = Ver
            frmMain.cVerTriggers.Visible = Ver
            frmMain.lListado(4).Visible = Ver
        Case 7
            frmMain.lstParticle.Visible = Ver
            frmMain.cmdAdd.Visible = Ver
            frmMain.cmdDel.Visible = Ver
            frmMain.Life.Visible = Ver
            frmMain.Label2.Visible = Ver
            frmMain.cmdView.Visible = Ver
        Case 8
            frmMain.cLuces.Visible = Ver
    End Select
    If Ver = True Then
        vMostrando = Numero
        If Numero < 0 Or Numero > 8 Then Exit Sub
        If frmMain.SelectPanel(Numero).value = False Then
            frmMain.SelectPanel(Numero).value = True
        End If
    Else
        If Numero < 0 Or Numero > 8 Then Exit Sub
        If frmMain.SelectPanel(Numero).value = True Then
            frmMain.SelectPanel(Numero).value = False
        End If
    End If
End Sub

''
' Filtra del Listado de Elementos de una Funcion
'
' @param Numero Indica la funcion a Filtrar

Public Sub Filtrar(ByVal Numero As Byte)
'*************************************************
'Author: ^[GS]^
'Last modified: 26/05/06
'*************************************************

    Dim vMaximo As Integer
    Dim vDatos As String
    Dim NumI As Integer
    Dim i As Integer
    Dim j As Integer
    
    If frmMain.cFiltro(Numero).ListCount > 5 Then
        frmMain.cFiltro(Numero).RemoveItem 0
    End If
    frmMain.cFiltro(Numero).AddItem frmMain.cFiltro(Numero).text
    frmMain.lListado(Numero).Clear
        
    Select Case Numero
        Case 0 ' superficie
            vMaximo = MaxSup
        Case 1 ' NPCs
            vMaximo = NumNPCs - 1
        Case 2 ' NPCs Hostiles
            'vMaximo = NumNPCsHOST - 1
        Case 3 ' Objetos
            vMaximo = NumOBJs - 1
    End Select
    
    For i = 0 To vMaximo
    
        Select Case Numero
            Case 0 ' superficie
                vDatos = SupData(i).name
                NumI = i
            Case 1 ' NPCs
                vDatos = NpcData(i + 1).name
                NumI = i + 1
            Case 2 ' NPCs Hostiles
                'vDatos = NpcData(i + 500).name
                'NumI = i + 500
            Case 3 ' Objetos
                vDatos = ObjData(i + 1).name
                NumI = i + 1
        End Select
        
        For j = 1 To Len(vDatos)
            If UCase$(mid$(vDatos & Str(i), j, Len(frmMain.cFiltro(Numero).text))) = UCase$(frmMain.cFiltro(Numero).text) Or LenB(frmMain.cFiltro(Numero).text) = 0 Then
                frmMain.lListado(Numero).AddItem vDatos & " - #" & NumI
                Exit For
            End If
        Next
    Next
End Sub

Public Function DameGrhIndex(ByVal GrhIn As Integer) As Integer
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************

DameGrhIndex = SupData(GrhIn).Grh

If SupData(GrhIn).Width > 0 Then
    frmConfigSup.MOSAICO.value = vbChecked
    frmConfigSup.mAncho.text = SupData(GrhIn).Width
    frmConfigSup.mLargo.text = SupData(GrhIn).Height
Else
    frmConfigSup.MOSAICO.value = vbUnchecked
    frmConfigSup.mAncho.text = "0"
    frmConfigSup.mLargo.text = "0"
End If



End Function

Public Sub fPreviewGrh(ByVal GrhIn As Integer)
'*************************************************
'Author: Unkwown
'Last modified: 22/05/06
'*************************************************

If Val(GrhIn) < 1 Then
  frmMain.cGrh.text = MaxGrhs
  Exit Sub
End If

If Val(GrhIn) > MaxGrhs Then
  frmMain.cGrh.text = 1
  Exit Sub
End If

'Change CurrentGrh
CurrentGrh.GrhIndex = GrhIn
CurrentGrh.Started = 1
CurrentGrh.FrameCounter = 1
CurrentGrh.speed = GrhData(CurrentGrh.GrhIndex).speed

End Sub

''
' Indica la accion de mostrar Vista Previa de la Superficie seleccionada
'

Public Sub VistaPreviaDeSup()
'*************************************************
'Author: ^[GS]^
'Last modified: 26/05/06
'*************************************************
Dim SR As RECT, DR As RECT

D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0

If CurrentGrh.GrhIndex = 0 Then Exit Sub
    If frmConfigSup.MOSAICO = vbUnchecked Then
        DR.Left = 0
        DR.Top = 0
        DR.Bottom = (GrhData(CurrentGrh.GrhIndex).pixelHeight)
        DR.Right = (GrhData(CurrentGrh.GrhIndex).pixelWidth)
        SR.Left = GrhData(CurrentGrh.GrhIndex).SX
        SR.Top = GrhData(CurrentGrh.GrhIndex).SY
        SR.Bottom = SR.Top + (GrhData(CurrentGrh.GrhIndex).pixelHeight)
        SR.Right = SR.Left + (GrhData(CurrentGrh.GrhIndex).pixelWidth)
        Call DrawGrhtoHdc(frmMain.PreviewGrh, CurrentGrh.GrhIndex, 1, 1)
    Else
        Dim X As Integer, y As Integer, j As Integer, i As Integer
        Dim Cont As Integer
        For i = 1 To CInt(Val(frmConfigSup.mLargo))
            For j = 1 To CInt(Val(frmConfigSup.mAncho))
                DR.Left = (j - 1) * 32
                DR.Top = (i - 1) * 32
                DR.Right = j * 32
                DR.Bottom = i * 32
                SR.Left = GrhData(CurrentGrh.GrhIndex).SX
                SR.Top = GrhData(CurrentGrh.GrhIndex).SY
                SR.Right = SR.Left + GrhData(CurrentGrh.GrhIndex).pixelWidth
                SR.Bottom = SR.Top + GrhData(CurrentGrh.GrhIndex).pixelHeight
                Call DrawGrhtoHdc(frmMain.PreviewGrh, CurrentGrh.GrhIndex, DR.Left, DR.Top)
                If Cont < CInt(Val(frmConfigSup.mLargo)) * CInt(Val(frmConfigSup.mAncho)) Then _
                    Cont = Cont + 1: CurrentGrh.GrhIndex = CurrentGrh.GrhIndex + 1
            Next
        Next
        CurrentGrh.GrhIndex = CurrentGrh.GrhIndex - Cont
    End If
End Sub
