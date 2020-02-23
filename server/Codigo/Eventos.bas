Attribute VB_Name = "Eventos"
'********************************Modulo Eventos*********************************
'Author: Matías Ignacio Rojo (MaxTus)
'Last Modification: 13/12/2011
'Generar eventos automáticos a patir de la variable TIEMPO.
'******************************************************************************

'******************************************************************************
'Determina la activación del evento
'******************************************************************************

Public Sub HappyHourAzar()
'Cada vez que se inicie una nueva hora, hay X probabilidad de que se active
'el evento.

Dim X As Byte

    If mid(Format(Time, "HH:MM:SS"), 4, 2) = 0 Then
        X = RandomNumber(1, 10)
        If X = 1 Then
            HappyHourAC = True
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos Automáticos> Se ha dado inicio el evento de " & _
                                                "Experiencia x2, el mismo finalizará a las " & mid(Format(Time, "HH:MM:SS"), 1, 2) + 1, FontTypeNames.FONTTYPE_GMMSG))
        End If
    End If
    
End Sub

