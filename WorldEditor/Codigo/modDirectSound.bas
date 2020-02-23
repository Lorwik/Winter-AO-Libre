Attribute VB_Name = "modDirectSound"
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
' modDirectSound
'
' @remarks Operaciones de control de Sonidos por DirectX
' @author unkwown
' @version 0.0.01
' @date 20060520

Public Sub IniciarDirectSound()
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
Err.Clear
On Error GoTo fin
    
    '<----------------Direct Music--------------->
    Set Perf = DirectX.DirectMusicPerformanceCreate()
    Call Perf.Init(Nothing, 0)
    Perf.SetPort -1, 80
    Call Perf.SetMasterAutoDownload(True)
    '<------------------------------------------->
    
    Set DirectSound = DirectX.DirectSoundCreate("")
    If Err Then
        MsgBox "Error iniciando DirectSound"
        End
    End If
    
    LastSoundBufferUsed = 1
    
    
    Exit Sub
fin:
End
End Sub

Public Sub LiberarDirectSound()
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
Dim cloop As Integer
For cloop = 1 To NumSoundBuffers
    Set DSBuffers(cloop) = Nothing
Next cloop
Set DirectSound = Nothing
End Sub
