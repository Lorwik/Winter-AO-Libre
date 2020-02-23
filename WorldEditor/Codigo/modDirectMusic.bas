Attribute VB_Name = "modDirectMusic"
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
' modDirectMusic
'
' @remarks Operaciones de control de MIDIs por DirectX
' @author unkwown
' @version 0.0.01
' @date 20060520

Option Explicit

Public Sub CargarMIDI(Archivo As String)
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
On Error GoTo fin
    
    If Loader Is Nothing Then Set Loader = DirectX.DirectMusicLoaderCreate()
    Set Seg = Loader.LoadSegment(Archivo)
    Set Loader = Nothing 'Liberamos el cargador
    Exit Sub
fin:
    MsgBox ("Error producido en 'CargarMIDI' " & Err.Description & " " & Err.Number & " " & Archivo)

End Sub

Public Sub Stop_Midi()
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
If IsPlayingCheck Then
     If Perf.IsPlaying(Seg, SegState) = True Then
            Call Perf.Stop(Seg, SegState, 0, 0)
     End If
     IsPlayingCheck = False
     Seg.SetStartPoint (0)
     Call Perf.Reset(0)
End If
End Sub

Public Sub Play_Midi()
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
On Error GoTo fin
        If IsPlayingCheck Then Stop_Midi
        If Perf.IsPlaying(Seg, SegState) = True Then
            Call Perf.Stop(Seg, SegState, 0, 0)
        End If
        Seg.SetStartPoint (0)
        Set SegState = Perf.PlaySegment(Seg, 0, 0)
        IsPlayingCheck = True
        Exit Sub
fin:
    MsgBox "Error producido en Public Sub Play_Midi()"

End Sub

Function Sonando()
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
Sonando = (Perf.IsPlaying(Seg, SegState) = True)
End Function




