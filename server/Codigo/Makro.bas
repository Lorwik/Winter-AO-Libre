Attribute VB_Name = "Makro"
'********************************Modulo Makro*********************************
'Author: Matías Ignacio Rojo (MaxTus)
'Last Modification: 02/12/2011
'Control asistido de trabajo.
'******************************************************************************

Option Explicit

'******************************************************************************
'Requisitos para pescar
'******************************************************************************

Public Function PuedePescar(ByVal UserIndex As Integer) As Boolean

    Dim DummyINT As Integer

    With UserList(UserIndex)
    
        DummyINT = .Invent.WeaponEqpObjIndex
                
        If DummyINT = 0 Then
            .flags.Makro = 0
            Call WriteConsoleMsg(UserIndex, "Necesitas una caña o una red para atrapar peces. Dejas de trabajar...", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(UserIndex, "*Macro de pesca asistido activado*", FontTypeNames.FONTTYPE_VENENO)
            PuedePescar = False
            Exit Function
        End If
        
        If .Stats.MinSta <= 5 Then
            Call WriteConsoleMsg(UserIndex, "Te encuentras demasiado cansado. Dejas de trabajar...", FontTypeNames.FONTTYPE_INFO)
            .flags.Makro = 0
            Call WriteConsoleMsg(UserIndex, "*Macro de pesca asistido desactivado*", FontTypeNames.FONTTYPE_VENENO)
            PuedePescar = False
            Exit Function
        End If
                
        If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = 1 Then
            Call WriteConsoleMsg(UserIndex, "No puedes pescar desde donde te encuentras.", FontTypeNames.FONTTYPE_INFO)
            .flags.Makro = 0
            Call WriteConsoleMsg(UserIndex, "*Macro de pesca asistido desactivado*", FontTypeNames.FONTTYPE_VENENO)
            PuedePescar = False
            Exit Function
        End If
        
        PuedePescar = True
        
    End With
    
End Function

'******************************************************************************
'Requisitos para Talar
'******************************************************************************

Public Function PuedeTalar(ByVal UserIndex As Integer) As Boolean

With UserList(UserIndex)

    If .flags.Equitando Then
        Call WriteConsoleMsg(UserIndex, "No puede trabajar estando montando. Dejas de trabajar...", FontTypeNames.FONTTYPE_INFO)
        .flags.Makro = 0
        Call WriteConsoleMsg(UserIndex, "*Macro de Leña asistido desactivado.*", FontTypeNames.FONTTYPE_VENENO)
        PuedeTalar = False
        Exit Function
    End If
                
    If MapInfo(UserList(UserIndex).Pos.map).Pk = False Then
        Call WriteConsoleMsg(UserIndex, "No puedes extraer leñas dentro de la ciudad. Dejas de trabajar...", FontTypeNames.FONTTYPE_INFO)
        .flags.Makro = 0
        Call WriteConsoleMsg(UserIndex, "*Macro de Leña asistido desactivado.*", FontTypeNames.FONTTYPE_VENENO)
        PuedeTalar = False
        Exit Function
    End If
                
    If .Invent.WeaponEqpObjIndex = 0 Then
        Call WriteConsoleMsg(UserIndex, "Deberías equiparte el hacha.", FontTypeNames.FONTTYPE_INFO)
        .flags.Makro = 0
        Call WriteConsoleMsg(UserIndex, "*Macro de Leña asistido desactivado.*", FontTypeNames.FONTTYPE_VENENO)
        PuedeTalar = False
        Exit Function
    End If
    
    If .Stats.MinSta <= 5 Then
            Call WriteConsoleMsg(UserIndex, "Te encuentras demasiado cansado. Dejas de trabajar...", FontTypeNames.FONTTYPE_INFO)
            .flags.Makro = 0
            Call WriteConsoleMsg(UserIndex, "*Macro de Leña asistido desactivado.*", FontTypeNames.FONTTYPE_VENENO)
            PuedeTalar = False
            Exit Function
        End If
                
    If .Invent.WeaponEqpObjIndex <> HACHA_LEÑADOR Then
        .flags.Makro = 0
        Call WriteConsoleMsg(UserIndex, "*Macro de Leña asistido desactivado.*", FontTypeNames.FONTTYPE_VENENO)
        PuedeTalar = False
        Exit Function
    End If
    
    PuedeTalar = True
                    
End With

End Function

'******************************************************************************
'Requisitos para minar
'******************************************************************************

Public Function PuedeMinar(ByVal UserIndex As Integer) As Boolean

With UserList(UserIndex)

    If UserList(UserIndex).flags.Equitando Then
        Call WriteConsoleMsg(UserIndex, "No puedes minar estando montado... Dejas de trabajar", FontTypeNames.FONTTYPE_INFO)
        .flags.Makro = 0
        Call WriteConsoleMsg(UserIndex, "*Macro de Mineria asistido desactivado.*", FontTypeNames.FONTTYPE_VENENO)
        PuedeMinar = False
        Exit Function
    End If
                
    If .Invent.WeaponEqpObjIndex = 0 Then
        .flags.Makro = 0
        Call WriteConsoleMsg(UserIndex, "*Macro de Mineria asistido desactivado.*", FontTypeNames.FONTTYPE_VENENO)
        PuedeMinar = False
        Exit Function
    End If
                
    If .Invent.WeaponEqpObjIndex <> PIQUETE_MINERO Then
        .flags.Makro = 0
        Call WriteConsoleMsg(UserIndex, "*Macro de Mineria asistido desactivado.*", FontTypeNames.FONTTYPE_VENENO)
        PuedeMinar = False
        Exit Function
    End If
    
    If .Stats.MinSta <= 5 Then
        Call WriteConsoleMsg(UserIndex, "Te encuentras demasiado cansado. Dejas de trabajar...", FontTypeNames.FONTTYPE_INFO)
        .flags.Makro = 0
        Call WriteConsoleMsg(UserIndex, "*Macro de Mineria asistido desactivado.*", FontTypeNames.FONTTYPE_VENENO)
        PuedeMinar = False
        Exit Function
    End If
    
    PuedeMinar = True
    
End With

End Function

'******************************************************************************
'Requisitos para lingotear
'******************************************************************************

Public Function PuedeLingotear(ByVal UserIndex As Integer) As Boolean

With UserList(UserIndex)

    If .flags.Equitando Then
        Call WriteConsoleMsg(UserIndex, "No puedes fundir minerales estando montado... Dejas de trabajar.", FontTypeNames.FONTTYPE_INFO)
        .flags.Makro = 0
        Call WriteConsoleMsg(UserIndex, "*Macro de Fundición asistido desactivado.*", FontTypeNames.FONTTYPE_VENENO)
        PuedeLingotear = False
        Exit Function
    End If
                
    'Check there is a proper item there
    If .flags.TargetObj > 0 Then
        If ObjData(.flags.TargetObj).OBJType = eOBJType.otFragua Then
        'Validate other items
            If .flags.TargetObjInvSlot < 1 Or .flags.TargetObjInvSlot > MAX_INVENTORY_SLOTS Then
                Call WriteConsoleMsg(UserIndex, "No tienes mas espacio en tu inventario... Dejas de trabajar.", FontTypeNames.FONTTYPE_INFO)
                .flags.Makro = 0
                Call WriteConsoleMsg(UserIndex, "*Macro de Fundición asistido desactivado.*", FontTypeNames.FONTTYPE_VENENO)
                PuedeLingotear = False
                Exit Function
            End If
                        
            ''chequeamos que no se zarpe duplicando oro
            If .Invent.Object(.flags.TargetObjInvSlot).ObjIndex <> .flags.TargetObjInvIndex Then
                If .Invent.Object(.flags.TargetObjInvSlot).ObjIndex = 0 Or .Invent.Object(.flags.TargetObjInvSlot).amount = 0 Then
                    .flags.Makro = 0
                    Call WriteConsoleMsg(UserIndex, "*Macro de Fundición asistido desactivado.*", FontTypeNames.FONTTYPE_VENENO)
                    PuedeLingotear = False
                    Exit Function
                End If
                            
                            ''FUISTE
                Call WriteErrorMsg(UserIndex, "Has sido expulsado por el sistema anti cheats.")
                Call FlushBuffer(UserIndex)
                Call CloseSocket(UserIndex)
                Exit Function
            End If
            
            'Puede trabajar ;)
            PuedeLingotear = True
            
        Else
            Call WriteConsoleMsg(UserIndex, "No hay ninguna fragua allí... Dejas de trabajar.", FontTypeNames.FONTTYPE_INFO)
            .flags.Makro = 0
            Call WriteConsoleMsg(UserIndex, "*Macro de Fundición asistido desactivado.*", FontTypeNames.FONTTYPE_VENENO)
            PuedeLingotear = False
            Exit Function
        End If
    Else
        Call WriteConsoleMsg(UserIndex, "No hay ninguna fragua allí... Dejas de trabajar.", FontTypeNames.FONTTYPE_INFO)
        .flags.Makro = 0
        Call WriteConsoleMsg(UserIndex, "*Macro de Fundición asistido desactivado.*", FontTypeNames.FONTTYPE_VENENO)
        PuedeLingotear = False
        Exit Function
    End If
            
End With

End Function

'******************************************************************************
'Inicia la actividad
'******************************************************************************

Public Sub MakroTrabajo(ByVal UserIndex As Integer, ByRef Tarea As eMakro)
        
        Select Case Tarea
            Case eMakro.Pescar
                If PuedePescar(UserIndex) Then
                    DoPescar UserIndex
                End If
                
            Case eMakro.Talar
                If PuedeTalar(UserIndex) Then
                    DoTalar UserIndex
                End If
                
            Case eMakro.Minar
                If PuedeMinar(UserIndex) Then
                    DoMineria UserIndex
                End If
                
            Case eMakro.PescarRed
                If PuedePescar(UserIndex) Then
                    DoPescarRed UserIndex
                End If
                
            Case eMakro.Lingotear
                If PuedeLingotear(UserIndex) Then
                    FundirMineral UserIndex
                End If
        End Select
        
End Sub
