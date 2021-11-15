VERSION 5.00
Begin VB.Form FrmMapa 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mapa del Mundo de Winter-AO"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmMapa.frx":0000
   ScaleHeight     =   8055
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image65 
      Height          =   375
      Left            =   2880
      Top             =   4440
      Width           =   735
   End
   Begin VB.Image Image60 
      Height          =   735
      Left            =   4920
      Top             =   960
      Width           =   615
   End
   Begin VB.Image Image58 
      Height          =   375
      Left            =   360
      Top             =   240
      Width           =   5895
   End
   Begin VB.Image Image61 
      Height          =   1455
      Left            =   2280
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Image Image62 
      Height          =   615
      Left            =   3000
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Image Image59 
      Height          =   4215
      Left            =   5520
      Top             =   240
      Width           =   735
   End
   Begin VB.Image Image64 
      Height          =   375
      Left            =   3600
      Top             =   4440
      Width           =   2655
   End
   Begin VB.Image Image63 
      Height          =   375
      Left            =   960
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Image Image57 
      Height          =   4575
      Left            =   240
      Top             =   240
      Width           =   735
   End
   Begin VB.Image Image56 
      Height          =   375
      Left            =   4920
      Top             =   600
      Width           =   615
   End
   Begin VB.Image Image55 
      Height          =   375
      Left            =   4200
      Top             =   600
      Width           =   735
   End
   Begin VB.Image Image54 
      Height          =   375
      Left            =   3600
      Top             =   600
      Width           =   615
   End
   Begin VB.Image Image53 
      Height          =   375
      Left            =   2880
      Top             =   600
      Width           =   735
   End
   Begin VB.Image Image52 
      Height          =   375
      Left            =   2280
      Top             =   600
      Width           =   615
   End
   Begin VB.Image Image51 
      Height          =   375
      Left            =   1560
      Top             =   600
      Width           =   615
   End
   Begin VB.Image Image50 
      Height          =   375
      Left            =   960
      Top             =   600
      Width           =   615
   End
   Begin VB.Image Image49 
      Height          =   375
      Left            =   960
      Top             =   960
      Width           =   615
   End
   Begin VB.Image Image48 
      Height          =   375
      Left            =   1560
      Top             =   960
      Width           =   735
   End
   Begin VB.Image Image47 
      Height          =   375
      Left            =   2280
      Top             =   960
      Width           =   615
   End
   Begin VB.Image Image46 
      Height          =   375
      Left            =   2880
      Top             =   960
      Width           =   735
   End
   Begin VB.Image Image45 
      Height          =   375
      Left            =   3600
      Top             =   960
      Width           =   615
   End
   Begin VB.Image Image44 
      Height          =   375
      Left            =   4200
      Top             =   960
      Width           =   735
   End
   Begin VB.Image Image43 
      Height          =   375
      Left            =   4200
      Top             =   1320
      Width           =   735
   End
   Begin VB.Image Image42 
      Height          =   375
      Left            =   3600
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image Image41 
      Height          =   375
      Left            =   2880
      Top             =   1320
      Width           =   735
   End
   Begin VB.Image Image40 
      Height          =   375
      Left            =   2280
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image Image39 
      Height          =   375
      Left            =   1560
      Top             =   1320
      Width           =   735
   End
   Begin VB.Image Image38 
      Height          =   375
      Left            =   960
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image Image37 
      Height          =   375
      Left            =   4920
      Top             =   1680
      Width           =   615
   End
   Begin VB.Image Image36 
      Height          =   375
      Left            =   4320
      Top             =   1680
      Width           =   615
   End
   Begin VB.Image Image35 
      Height          =   375
      Left            =   3600
      Top             =   1680
      Width           =   615
   End
   Begin VB.Image Image34 
      Height          =   375
      Left            =   2880
      Top             =   1680
      Width           =   735
   End
   Begin VB.Image Image33 
      Height          =   375
      Left            =   2280
      Top             =   1680
      Width           =   615
   End
   Begin VB.Image Image32 
      Height          =   375
      Left            =   1680
      Top             =   1680
      Width           =   495
   End
   Begin VB.Image Image31 
      Height          =   375
      Left            =   960
      Top             =   2040
      Width           =   615
   End
   Begin VB.Image Image30 
      Height          =   255
      Left            =   1560
      Top             =   2040
      Width           =   735
   End
   Begin VB.Image Image29 
      Height          =   255
      Left            =   2280
      Top             =   2040
      Width           =   615
   End
   Begin VB.Image Image28 
      Height          =   375
      Left            =   2880
      Top             =   2040
      Width           =   735
   End
   Begin VB.Image Image27 
      Height          =   255
      Left            =   3600
      Top             =   2040
      Width           =   615
   End
   Begin VB.Image Image26 
      Height          =   375
      Left            =   4320
      Top             =   2040
      Width           =   615
   End
   Begin VB.Image Image25 
      Height          =   255
      Left            =   4320
      Top             =   2400
      Width           =   495
   End
   Begin VB.Image Image24 
      Height          =   255
      Left            =   4920
      Top             =   2040
      Width           =   615
   End
   Begin VB.Image Image23 
      Height          =   255
      Left            =   4920
      Top             =   2400
      Width           =   615
   End
   Begin VB.Image Image22 
      Height          =   255
      Left            =   3600
      Top             =   2400
      Width           =   615
   End
   Begin VB.Image Image21 
      Height          =   255
      Left            =   2880
      Top             =   2400
      Width           =   735
   End
   Begin VB.Image Image20 
      Height          =   255
      Left            =   2280
      Top             =   2400
      Width           =   615
   End
   Begin VB.Image Image19 
      Height          =   255
      Left            =   1560
      Top             =   2400
      Width           =   735
   End
   Begin VB.Image Image18 
      Height          =   255
      Left            =   960
      Top             =   2400
      Width           =   615
   End
   Begin VB.Image Image17 
      Height          =   375
      Left            =   2280
      Top             =   2640
      Width           =   615
   End
   Begin VB.Image Image16 
      Height          =   255
      Left            =   960
      Top             =   2760
      Width           =   615
   End
   Begin VB.Image Image15 
      Height          =   375
      Left            =   1560
      Top             =   2640
      Width           =   735
   End
   Begin VB.Image Image14 
      Height          =   375
      Left            =   960
      Top             =   3000
      Width           =   615
   End
   Begin VB.Image Image13 
      Height          =   375
      Left            =   1560
      Top             =   3000
      Width           =   735
   End
   Begin VB.Image Image12 
      Height          =   375
      Left            =   960
      Top             =   3360
      Width           =   615
   End
   Begin VB.Image Image11 
      Height          =   375
      Left            =   1560
      Top             =   3360
      Width           =   735
   End
   Begin VB.Image Image10 
      Height          =   375
      Left            =   1560
      Top             =   3720
      Width           =   735
   End
   Begin VB.Image Image9 
      Height          =   375
      Left            =   960
      Top             =   3720
      Width           =   615
   End
   Begin VB.Image Image8 
      Height          =   375
      Left            =   960
      Top             =   4080
      Width           =   615
   End
   Begin VB.Image Image7 
      Height          =   375
      Left            =   240
      Top             =   7440
      Width           =   735
   End
   Begin VB.Image Image6 
      Height          =   375
      Left            =   240
      Top             =   6960
      Width           =   735
   End
   Begin VB.Image Image5 
      Height          =   375
      Left            =   360
      Top             =   6480
      Width           =   495
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   240
      Top             =   6000
      Width           =   735
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   240
      Top             =   5520
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   240
      Top             =   5030
      Width           =   735
   End
   Begin VB.Label Nombre 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   5595
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   1560
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label informacion 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Pinche sobre un mapa para obtener informacion acerca de el."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1095
      Left            =   2640
      TabIndex        =   0
      Top             =   6480
      Width           =   3255
   End
End
Attribute VB_Name = "FrmMapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Picture = General_Load_Picture_From_Resource("MapadelMundo.gif")
End Sub

Private Sub Image1_Click()
Nombre.Caption = "Ramx"
informacion.Caption = "Ramx es un pequeño y humilde pueblo de pescadores. En este pequeño pueblo es donde todos los nuevos y curiosos aventureros emprenden su viaje hacia lo desconocido. Entrada de las alcantarillas."
End Sub

Private Sub Image10_Click()
Nombre.Caption = "Afueras de Ramx"
informacion.Caption = "En este bosque encontraras Lobos, murciélagos, serpientes y Cocodrilo Joven."
End Sub

Private Sub Image11_Click()
Nombre.Caption = "Bosques del Sur"
informacion.Caption = "En este bosque encontraras Cocodrilo Joven, Cocodrilo Viejo, Esqueletos y Lobos."
End Sub

Private Sub Image12_Click()
Nombre.Caption = "Bosques del Sur"
informacion.Caption = "En este bosque encontraras Equeletos, goblin y serpientes."
End Sub

Private Sub Image13_Click()
Nombre.Caption = "Bosques del Sur"
informacion.Caption = "En este bosque encontraras Esqueletos, Zombis y Goblin."
End Sub

Private Sub Image14_Click()
Nombre.Caption = "Bosques del Sur"
informacion.Caption = "En este bosque encontraras Lobos, murciélagos y Gallos Salvajes."
End Sub

Private Sub Image15_Click()
Nombre.Caption = "Bosques Oscuros"
informacion.Caption = "En este bosque encontraras Lord Zombis, Zombis y Arañas Gigantes."
End Sub

Private Sub Image16_Click()
Nombre.Caption = "Bosques del Sur"
informacion.Caption = "En este bosque encontraras Esqueletos, Lobos, Globin y Serpientes."
End Sub

Private Sub Image17_Click()
Nombre.Caption = "Bosques Oscuros"
informacion.Caption = "En este bosque encontraras Arañas Gigantes, Zombis y Esqueletos."
End Sub

Private Sub Image18_Click()
Nombre.Caption = "Bosques Oscuros"
informacion.Caption = "En este bosque encontraras Arañas Gigantes, Lord Zombis y Esqueletos."
End Sub

Private Sub Image19_Click()
Nombre.Caption = "Entrada Dungeon Spectral"
informacion.Caption = "Entrada del Dungeon Spectral, donde encontratas Arañas Gigantes, Lord Zombis, Esqueletos y Gladiador."
End Sub

Private Sub Image2_Click()
Nombre.Caption = "Agua"
informacion.Caption = "Todos los cuadrados que son así, significa que es agua."
End Sub

Private Sub Image20_Click()
Nombre.Caption = "Bosques Oscuros"
informacion.Caption = "Aqui encontraras Orcos, Oso Pardo y Globin."
End Sub

Private Sub Image21_Click()
Nombre.Caption = "Bosques Sagrados"
informacion.Caption = "En este bosque podras encontrar Oso Pardo, Tortugas Gigantes y Gallos Salvajes."
End Sub

Private Sub Image22_Click()
Nombre.Caption = "Bosques Sagrados"
informacion.Caption = "En este bosque podras encontrar Oso Pardo, Tortugas Gigantes y Gallos Salvajes."
End Sub

Private Sub Image23_Click()
Nombre.Caption = "Bosques Sagrados"
informacion.Caption = "En este bosque podras encontrar Oso Pardo, Tortugas Gigantes y Gallos Salvajes."
End Sub

Private Sub Image24_Click()
Nombre.Caption = "Castillo Oeste"
informacion.Caption = "Aqui se encuentra el Rey del Castillo Oeste, aquel clan que lo mate conquistara dicho castillo."
End Sub

Private Sub Image26_Click()
Nombre.Caption = "Bosques Abiertos"
informacion.Caption = "Aqui encontraras Lobos y Gladiadores."
End Sub

Private Sub Image27_Click()
Nombre.Caption = "Afueras de Winderbill"
informacion.Caption = "Aqui encontraras Lobos y Asesinos."
End Sub

Private Sub Image28_Click()
Nombre.Caption = "Afueras de Winderbill"
informacion.Caption = "Aqui encontraras Asesinos y muercielagos."
End Sub

Private Sub Image29_Click()
Nombre.Caption = "Afueras de Winderbill"
informacion.Caption = "Aqui encontraras Arañas Gigantes, Jefe Orcos y Huargos."
End Sub

Private Sub Image3_Click()
Nombre.Caption = "Bosques"
informacion.Caption = "Todos los cuadrados que son así, significa que son bosques."
End Sub

Private Sub Image30_Click()
Nombre.Caption = "Bosques Oscuros"
informacion.Caption = "Aqui encontraras Arañas Gigantes, Orcos, Jefe Orcos, Orcos Brujo."
End Sub

Private Sub Image31_Click()
Nombre.Caption = "Bosques Oscuros"
informacion.Caption = "Aqui encontraras Arañas Gigantes, Lobos y Serpiente."
End Sub

Private Sub Image32_Click()
Nombre.Caption = "Bosques Oscuros"
informacion.Caption = "Aqui encontraras asesinos, tortugas gigantes y lobos."
End Sub

Private Sub Image33_Click()
Nombre.Caption = "Bosques Oscuros"
informacion.Caption = "Aqui encontraras duendes magicos, Zombis, Lobos y Serpientes."
End Sub

Private Sub Image34_Click()
Nombre.Caption = "Winderbill"
informacion.Caption = "Capital de la Armada Real, ciudad prospera desde siempre."
End Sub

Private Sub Image35_Click()
Nombre.Caption = "Afuertas de Winderbill"
informacion.Caption = "Aqui encontraras Lobos, gallos salvajes y esqueletos."
End Sub

Private Sub Image36_Click()
Nombre.Caption = "Bosque Abiertos"
informacion.Caption = "Aqui encontraras Esqueleto Guerrero, Bandidos, Agilas."
End Sub

Private Sub Image37_Click()
Nombre.Caption = "Costa Este"
informacion.Caption = "Aqui encontraras Esqueletos Guerreros y Lobos."
End Sub

Private Sub Image38_Click()
Nombre.Caption = "Bosques Oscuros"
informacion.Caption = "Aqui encontraras Arañas Gigantes, Lobos, Bandidos."
End Sub

Private Sub Image39_Click()
Nombre.Caption = "Bosques Oscuros"
informacion.Caption = "Aqui encontraras Bandidos, Asesinos y Lobos."
End Sub

Private Sub Image4_Click()
Nombre.Caption = "Ciudad"
informacion.Caption = "Todos los cuados que son así, significan que son ciudades."
End Sub

Private Sub Image40_Click()
Nombre.Caption = "Afueras de Winderbill"
informacion.Caption = "Aqui encontraras Murcielagos, Lobos y Bandidos."
End Sub

Private Sub Image41_Click()
Nombre.Caption = "Afueras de Winderbill"
informacion.Caption = "Aqui encontraras Murcielagos, Lobos y Bandidos."
End Sub

Private Sub Image42_Click()
Nombre.Caption = "Afueras de Winderbill"
informacion.Caption = "Aqui encontraras Murcielagos, Lobos y Serpientes."
End Sub

Private Sub Image43_Click()
Nombre.Caption = "Desierto"
informacion.Caption = "Aqui encontraras Arañas, escorpion y Serpientes."
End Sub

Private Sub Image44_Click()
Nombre.Caption = "Desierto"
informacion.Caption = "Aqui encontraras Arañas, escopion y Serpientes."
End Sub

Private Sub Image45_Click()
Nombre.Caption = "Desierto"
informacion.Caption = "Aqui encontraras Arañas, escopion y Serpientes."
End Sub

Private Sub Image46_Click()
Nombre.Caption = "Desierto"
informacion.Caption = "Aqui encontraras Arañas Gigantes, escopion y Serpientes."
End Sub

Private Sub Image47_Click()
Nombre.Caption = "Afueras de Kripus"
informacion.Caption = "Aqui encontraras Arañas Gigantes y Huargos."
End Sub

Private Sub Image48_Click()
Nombre.Caption = "Afueras de Kripus"
informacion.Caption = "Aqui encontraras Lobos y Huargos."
End Sub

Private Sub Image49_Click()
Nombre.Caption = "Afueras de Kripus"
informacion.Caption = "Aqui encontraras Lobos y Asesinos."
End Sub

Private Sub Image5_Click()
Nombre.Caption = "Dungeon"
informacion.Caption = "Todos los cuadrados que contenga este simbolo, significa que son dungeones."
End Sub

Private Sub Image50_Click()
Nombre.Caption = "Kripus"
informacion.Caption = "Ciudad donde reina el caos y el mal."
End Sub

Private Sub Image51_Click()
Nombre.Caption = "Bosques de Kripus"
informacion.Caption = "Aqui encontraras Arañas Gigantes, lobos y Huargos."
End Sub

Private Sub Image52_Click()
Nombre.Caption = "Bosques de Kripus"
informacion.Caption = "Aqui encontraras Arañas Gigantes y lobos."
End Sub

Private Sub Image53_Click()
Nombre.Caption = "Desierto"
informacion.Caption = "Aqui encontraras Arañas, escorpion y Serpientes."
End Sub

Private Sub Image54_Click()
Nombre.Caption = "Desierto"
informacion.Caption = "Aqui encontraras Arañas, escorpion y Serpientes."
End Sub

Private Sub Image55_Click()
Nombre.Caption = "Desierto"
informacion.Caption = "Aqui encontraras Arañas, escorpion y Serpientes."
End Sub

Private Sub Image56_Click()
Nombre.Caption = "Feinur"
informacion.Caption = "Ciudad construida por aquellos viajeros perdidos en el desierto.."
End Sub

Private Sub Image57_Click()
Nombre.Caption = "Mar"
informacion.Caption = "Sin informacion."
End Sub

Private Sub Image58_Click()
Nombre.Caption = "Mar"
informacion.Caption = "Sin informacion."
End Sub

Private Sub Image59_Click()
Nombre.Caption = "Mar"
informacion.Caption = "Sin informacion."
End Sub

Private Sub Image6_Click()
Nombre.Caption = "Desierto"
informacion.Caption = "Todos los cuadrados que son así, significa que es desierto."
End Sub

Private Sub Image60_Click()
Nombre.Caption = "Mar"
informacion.Caption = "Sin informacion."
End Sub

Private Sub Image61_Click()
Nombre.Caption = "Mar"
informacion.Caption = "Sin informacion."
End Sub

Private Sub Image62_Click()
Nombre.Caption = "Mar"
informacion.Caption = "Sin informacion."
End Sub

Private Sub Image63_Click()
Nombre.Caption = "Mar"
informacion.Caption = "Sin informacion."
End Sub

Private Sub Image64_Click()
Nombre.Caption = "Mar"
informacion.Caption = "Sin informacion."
End Sub

Private Sub Image65_Click()
Nombre.Caption = "Shakoud"
informacion.Caption = "Isla al sur del mundo de Winter, donde habian los seres mas pequeños de estas tierras."
End Sub

Private Sub Image7_Click()
Nombre.Caption = "Nieve"
informacion.Caption = "Todos los cuadrados que son así, significa que es nieve."
End Sub

Private Sub Image8_Click()
Nombre.Caption = "Afueras de Ramx"
informacion.Caption = "En este bosque encontraras Lobos, murciélagos, serpientes y Queaker, es un buen lugar de entrenamiento para aquellos aventureros que comiencen su largo viaje."
End Sub

Private Sub Image9_Click()
Nombre.Caption = "Bosques del Sur"
informacion.Caption = "En este bosque encontraras Lobos, murciélagos y serpientes."
End Sub
