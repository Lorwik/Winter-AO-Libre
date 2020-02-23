VERSION 5.00
Begin VB.Form frmCalculator 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1740
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   1740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtVida 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   240
      TabIndex        =   6
      Text            =   "0"
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtVidaInic 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Text            =   "0"
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtELV 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Text            =   "0"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.ComboBox cboRAZA 
      Height          =   315
      ItemData        =   "frmCalculator.frx":0000
      Left            =   240
      List            =   "frmCalculator.frx":0013
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.ComboBox cboCLASE 
      Height          =   315
      ItemData        =   "frmCalculator.frx":0040
      Left            =   240
      List            =   "frmCalculator.frx":0068
      TabIndex        =   2
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton cmdCalcular 
      Caption         =   "Calcular"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Clase:"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Raza:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Nivel:"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Vida inicial:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Vida actual:"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmcalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Clase As String
Dim Raza As String
Dim intVida As Integer
Dim intVidainic As Integer
Dim VidamenosVida As Integer
Dim intELV As Integer
Dim Promedio As Long




Private Sub cmdCalcular_Click()
Inicio:
Raza = cboRAZA
Clase = cboCLASE
intVida = txtVida.Text
intVidainic = txtVidaInic.Text
intELV = txtELV.Text
VidamenosVida = intVida - intVidainic
Promedio = VidamenosVida / (intELV - 1)

If intVida <= 0 Then
    MsgBox ("Error en VIDA")
    GoTo Inicio
End If
If intVidainic <= 0 Then
    MsgBox ("Error en VIDA INICIAL")
    GoTo Inicio
End If

If intELV <= 0 Then
    MsgBox ("Error en NIVEL")
    GoTo Inicio
End If
    


 MsgBox "Su Promedio es de " & Promedio, vbInformation, "Winter AO Return"
 
If UCase(Clase) = "GUERRERO" Or UCase(Clase) = "LADRON" Then
   If UCase(Raza) = "ENANO" Then
        
        If Promedio <= 9.9 And Promedio > 9.7 Then
            MsgBox ("¡Por los pelos!")
        ElseIf Promedio <= 9.7 Then
            MsgBox ("Lo siento, pero estas bajo del promedio")
        ElseIf Promedio = 10 Then
            MsgBox ("Estas en el promedio")
        ElseIf Promedio > 10 Then
            MsgBox ("¡Felicidades estas por encima del promedio!")

   End If
    End If
   If UCase(Raza) = "HUMANO" Then
        
        If Promedio <= 9.4 And Promedio > 9.2 Then
            MsgBox ("¡Por los pelos!")
        ElseIf Promedio <= 9.2 Then
            MsgBox ("Lo siento, pero estas bajo del promedio")
        ElseIf Promedio = 9.5 Then
            MsgBox ("Estas en el promedio")
        ElseIf Promedio > 9.5 Then
            MsgBox ("¡Felicidades estas por encima del promedio!")
            
    End If
        End If
   
   If UCase(Raza) = "ELFO OSCURO" Or UCase(Raza) = "GNOMO" Then
       
       If Promedio <= 8.4 And Promedio > 8.3 Then
            MsgBox ("¡Por los pelos!")
        ElseIf Promedio <= 8.3 Then
            MsgBox ("Lo siento, pero estas bajo del promedio")
        ElseIf Promedio = 8.5 Then
            MsgBox ("Estas en el promedio")
        ElseIf Promedio > 8.5 Then
            MsgBox ("¡Felicidades estas por encima del promedio!")

    End If
        End If
    
        
   If UCase(Raza) = "ELFO" Then
        
        If Promedio <= 8.9 And Promedio > 8.7 Then
            MsgBox ("¡Por los pelos!")
        ElseIf Promedio <= 8.7 Then
            MsgBox ("Lo siento, pero estas bajo del promedio")
        ElseIf Promedio = 9 Then
            MsgBox ("Estas en el promedio")
        ElseIf Promedio > 9 Then
            MsgBox ("¡Felicidades estas por encima del promedio!")
    End If
        End If
End If
    
If UCase(Clase) = "PALADIN" Or UCase(Clase) = "BANDIDO" Or UCase(Clase) = "CAZADOR" Or UCase(Clase) = "PIRATA" Then
    
    If UCase(Raza) = "ENANO" Then
         
        
        If Promedio <= 9.4 And Promedio > 9.3 Then
            MsgBox ("¡Por los pelos!")
        ElseIf Promedio <= 9.3 Then
            MsgBox ("Lo siento, pero estas bajo del promedio")
        ElseIf Promedio = 9.5 Then
            MsgBox ("Estas en el promedio")
        ElseIf Promedio > 9.5 Then
            MsgBox ("¡Felicidades estas por encima del promedio!")
        

    End If
        End If

   If UCase(Raza) = "HUMANO" Then
        
        If Promedio <= 8.9 And Promedio > 8.8 Then
            MsgBox ("¡Por los pelos!")
        ElseIf Promedio <= 8.8 Then
            MsgBox ("Lo siento, pero estas bajo del promedio")
        ElseIf Promedio = 9 Then
            MsgBox ("Estas en el promedio")
        ElseIf Promedio > 9 Then
            MsgBox ("¡Felicidades estas por encima del promedio!")
            
    End If
        End If
   
   If UCase(Raza) = "ELFO OSCURO" Or UCase(Raza) = "GNOMO" Then
       
       If Promedio <= 7.9 And Promedio > 7.8 Then
            MsgBox ("¡Por los pelos!")
        ElseIf Promedio <= 7.8 Then
            MsgBox ("Lo siento, pero estas bajo del promedio")
        ElseIf Promedio = 8 Then
            MsgBox ("Estas en el promedio")
        ElseIf Promedio > 8 Then
            MsgBox ("¡Felicidades estas por encima del promedio!")

    End If
        End If
    
        
   If UCase(Raza) = "ELFO" Then
        
        If Promedio <= 8.4 And Promedio > 8.3 Then
            MsgBox ("¡Por los pelos!")
        ElseIf Promedio <= 8.3 Then
            MsgBox ("Lo siento, pero estas bajo del promedio")
        ElseIf Promedio = 8.5 Then
            MsgBox ("Estas en el promedio")
        ElseIf Promedio > 8.5 Then
            MsgBox ("¡Felicidades estas por encima del promedio!!")
    End If
        End If
End If

If UCase(Clase) = "BARDO" Or UCase(Clase) = "CLERIGO" Or UCase(Clase) = "DRUIDA" Or UCase(Clase) = "ASESINO" Then
    
    If UCase(Raza) = "ENANO" Then
        
        If Promedio <= 8.4 And Promedio > 8.3 Then
            MsgBox ("¡Por los pelos!")
        ElseIf Promedio <= 8.3 Then
            MsgBox ("Lo siento, pero estas bajo del promedio")
        ElseIf Promedio = 8.5 Then
            MsgBox ("Estas en el promedio")
        ElseIf Promedio > 8.5 Then
            MsgBox ("¡Felicidades estas por encima del promedio!")
End If
   End If

   If UCase(Raza) = "HUMANO" Then
        
        If Promedio <= 7.9 And Promedio > 7.8 Then
            MsgBox ("Estas en el borde")
        ElseIf Promedio <= 7.8 Then
            MsgBox ("Lo siento, pero estas bajo del promedio")
        ElseIf Promedio = 8 Then
            MsgBox ("Estas en el promedio")
        ElseIf Promedio > 8 Then
            MsgBox ("¡Felicidades estas por encima del promedio!")
            
    End If
        End If
   
   If UCase(Raza) = "ELFO OSCURO" Or UCase(Raza) = "GNOMO" Then
       
       If Promedio <= 6.9 And Promedio > 6.8 Then
            MsgBox ("¡Por los pelos!")
        ElseIf Promedio <= 6.8 Then
            MsgBox ("Lo siento, pero estas bajo del promedio")
        ElseIf Promedio = 7 Then
            MsgBox ("Estas en el promedio")
        ElseIf Promedio > 7 Then
            MsgBox ("¡Felicidades estas por encima del promedio!")

    End If
        End If
    
        
   If UCase(Raza) = "ELFO" Then
        
        If Promedio <= 7.4 And Promedio > 7.3 Then
            MsgBox ("¡Por los pelos!")
        ElseIf Promedio <= 7.3 Then
            MsgBox ("Lo siento, pero estas bajo del promedio")
        ElseIf Promedio = 7.5 Then
            MsgBox ("Estas en el promedio")
        ElseIf Promedio > 7.5 Then
            MsgBox ("¡Felicidades estas por encima del promedio!")
    End If
        End If
End If
    

If UCase(Clase) = "MAGO" Then
    
    If UCase(Raza) = "ENANO" Then
        
        If Promedio <= 7.4 And Promedio > 7.3 Then
            MsgBox ("¡Por los pelos!")
        ElseIf Promedio <= 7.3 Then
            MsgBox ("Lo siento, pero estas bajo del promedio")
        ElseIf Promedio = 7.5 Then
            MsgBox ("Estas en el promedio")
        ElseIf Promedio > 7.5 Then
            MsgBox ("¡Felicidades estas por encima del promedio!")
End If
   End If

   If UCase(Raza) = "HUMANO" Then
        
        If Promedio <= 6.9 And Promedio > 6.8 Then
            MsgBox ("¡Por los pelos!")
        ElseIf Promedio <= 6.8 Then
            MsgBox ("Lo siento, pero estas bajo del promedio")
        ElseIf Promedio = 7 Then
            MsgBox ("Estas en el promedio")
        ElseIf Promedio > 7 Then
            MsgBox ("¡Felicidades estas por encima del promedio!")
            
    End If
        End If
   
   If UCase(Raza) = "ELFO OSCURO" Or UCase(Raza) = "GNOMO" Then
       
       If Promedio <= 5.9 And Promedio > 5.8 Then
            MsgBox ("¡Por los pelos!")
        ElseIf Promedio <= 5.8 Then
            MsgBox ("Lo siento, pero estas bajo del promedio")
        ElseIf Promedio = 6 Then
            MsgBox ("Estas en el promedio")
        ElseIf Promedio > 6 Then
            MsgBox ("¡Felicidades estas por encima del promedio!")

    End If
        End If
    
        
   If UCase(Raza) = "ELFO" Then
        
        If Promedio <= 6.4 And Promedio > 6.3 Then
            MsgBox ("¡Por los pelos!")
        ElseIf Promedio <= 6.3 Then
            MsgBox ("Lo siento, pero estas bajo del promedio")
        ElseIf Promedio = 6.5 Then
            MsgBox ("Estas en el promedio")
        ElseIf Promedio > 6.5 Then
            MsgBox ("¡Felicidades estas por encima del promedio!")
    End If

End If
End If

End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub
