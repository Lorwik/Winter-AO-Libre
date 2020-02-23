VERSION 5.00
Begin VB.Form frmCargando 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   7590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10050
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   506
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Consejo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Consejo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   5520
      Width           =   9495
   End
   Begin VB.Label Estado 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3600
      TabIndex        =   0
      Top             =   6240
      Width           =   3135
   End
   Begin VB.Image imgProgress 
      Height          =   495
      Left            =   2310
      Top             =   6090
      Width           =   5400
   End
End
Attribute VB_Name = "frmCargando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Private porcentajeActual As Integer
 
Private Const PROGRESS_DELAY = 10
Private Const PROGRESS_DELAY_BACKWARDS = 4
Private Const DEFAULT_PROGRESS_WIDTH = 336
Private Const DEFAULT_STEP_FORWARD = 1
Private Const DEFAULT_STEP_BACKWARDS = -3
 
Public Sub progresoConDelay(ByVal porcentaje As Integer)
 
If porcentaje = porcentajeActual Then Exit Sub
 
Dim step As Integer, stepInterval As Integer, Timer As Long, tickCount As Long
 
If (porcentaje > porcentajeActual) Then
    step = DEFAULT_STEP_FORWARD
    stepInterval = PROGRESS_DELAY
Else
    step = DEFAULT_STEP_BACKWARDS
    stepInterval = PROGRESS_DELAY_BACKWARDS
End If
 
Do Until compararPorcentaje(porcentaje, porcentajeActual, step)
    Do Until (Timer + stepInterval) <= GetTickCount()
        DoEvents
    Loop
    Timer = GetTickCount()
    porcentajeActual = porcentajeActual + step
    Call establecerProgreso(porcentajeActual)
Loop
 
End Sub
 
 
Public Sub establecerProgreso(ByVal nuevoPorcentaje As Integer)
 
If nuevoPorcentaje >= 0 And nuevoPorcentaje <= 100 Then
    imgProgress.Width = DEFAULT_PROGRESS_WIDTH * CLng(nuevoPorcentaje) / 100
ElseIf nuevoPorcentaje > 100 Then
    imgProgress.Width = DEFAULT_PROGRESS_WIDTH
Else
    imgProgress.Width = 0
End If
porcentajeActual = nuevoPorcentaje
 
End Sub
 
Private Function compararPorcentaje(ByVal porcentajeTarget As Integer, ByVal porcentajeAct As Integer, ByVal step As Integer) As Boolean
 
If step = DEFAULT_STEP_FORWARD Then
    compararPorcentaje = (porcentajeAct >= porcentajeTarget)
Else
    compararPorcentaje = (porcentajeAct <= porcentajeTarget)
End If
 
End Function

Private Sub Form_Load()
Me.Caption = Form_Caption
Me.Picture = General_Load_Picture_From_Resource("38.gif")
imgProgress.Picture = General_Load_Picture_From_Resource("113.gif")
Call MsgConsjo(RandomNumber(0, 16))
End Sub

Private Sub MsgConsjo(Index As Integer)
    Select Case Index
        Case 0
            Consejo.Caption = "Al hacer Quest, ¡obtendrás más experiencia, oro e ítems!"
        Case 1
            Consejo.Caption = "Las criaturas no siempre dropearan ítems."
        Case 2
            Consejo.Caption = "Puedes configurar los macros de comandos, en la configuración de teclas, del menú opciones."
        Case 3
            Consejo.Caption = "Antes de enviar una consulta a los GM, consulta el Manual."
        Case 4
            Consejo.Caption = "Los skills son mas fáciles de subir mientras seas Newbie."
        Case 5
            Consejo.Caption = "Si escribes ""/RANK"", podrás ver los usuarios con mas nivel, oro y muertes del server."
        Case 6
            Consejo.Caption = "Escribiendo el comando ""/TOP5"" podrás ver los 5 mejores usuarios del server."
        Case 7
            Consejo.Caption = "Si utilizas monturas obtendras ventajas como velocidad, refuerzo de la armadura y sigilo."
        Case 8
            Consejo.Caption = "Recomendamos jugar con todas las opciones de sonido activadas para disfrutar de una mayor experiencia."
        Case 9
            Consejo.Caption = "En cada ciudad encontraras equipos/magias diferente, en las capitales encontraras los mejores equipos a la venta."
        Case 10
            Consejo.Caption = "Los Titanes no pueden ser paralizados. Se recomienda el uso de un guerrero junto sus hermanos para derrotarlos."
        Case 11
            Consejo.Caption = "Debes de comprender que los GM's atienden muchas consultas, debes de esperar pacientemente tu turno."
        Case 12
            Consejo.Caption = "Los GM's son personas ocupadas que da soporte a todos los usuarios. ¡No les pidas objetos! ¡No te lo daran!"
        Case 13
            Consejo.Caption = "Para fundar un clan, necesitaras forjar el Anillo de la Hermandad, realizando una quest que se encuentra en Ramx."
        Case 14
            Consejo.Caption = "Si no encuentras ningun GM online o tu problema no puede ser resuelto online, te recomendamos visitar el foro."
        Case 15
            Consejo.Caption = "Apretando la letra ""Q"" podras ver el mapa del mundo."
        Case 16
            Consejo.Caption = "Los puntos de canjes se pueden conseguir por eventos oficiales o donaciones. Encontraras mas informacion en la web."
    End Select
End Sub
