VERSION 5.00
Begin VB.Form frmFractal 
   BackColor       =   &H00000000&
   Caption         =   "FRACTALES"
   ClientHeight    =   10500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13875
   LinkTopic       =   "Form1"
   ScaleHeight     =   10101.01
   ScaleMode       =   0  'User
   ScaleWidth      =   12862.11
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPaisaje 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PAISAJE ALEATORIO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdDragon 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DRAGON"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdCurva 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CURVA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdCuadros 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CUADROS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdCopoNieve 
      BackColor       =   &H00FFFFFF&
      Caption         =   "COPO NIEVE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmFractal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************************
'* PROYECTO      : FRACTALES
'* CONTENIDO     : PERMITE VISUALIZAR FORMAS FRACTALES
'* VERSION       : 1.1
'* AUTORES       : MIGUEL QUINTEIRO PIÑERO / MIGUEL QUINTEIRO FERNANDEZ
'* INICIO        : 04 DE MARZO DE 2014
'* ACTUALIZACION : 04 DE MARZO DE 2014
'****************************************************************************************
Option Explicit

' DECLARACION DE VARIABLES
Dim PI

' AL MOMENTO DE CARGAR EL FORMULARIO
Private Sub Form_Load()
  PI = 4 * Atn(1)
End Sub

' AL HACER CLICK SOBRE EL FORMULARIO
Sub Form_click()
  Cls
End Sub

' DIBUJA EL COPO DE NIEVE
Private Sub cmdCopoNieve_Click()
  Cls
  Scale (-100, 100)-(100, -100)
  Koch 0, 0, 120
End Sub

' SUBRUTINA PARA DIBUJAR EL COPO DE NIEVE
Sub Koch(Xpos, Ypos, Tamano)
  ReDim X(12), Y(12)
  Dim Desplazamiento As Single, i As Integer
  Dim Color
  'Color = QBColor(Int(15 * Rnd))
  Color = vbWhite
  If Tamano < 3 Then Exit Sub
  Desplazamiento = Tamano / 8
  X(1) = Xpos: Y(1) = Ypos + 4 * Desplazamiento
  X(2) = Xpos + Desplazamiento: Y(2) = Ypos + 2 * Desplazamiento
  X(3) = Xpos + 3 * Desplazamiento: Y(3) = Y(2)
  X(4) = Xpos + 2 * Desplazamiento: Y(4) = Ypos
  X(5) = X(3): Y(5) = Ypos - 2 * Desplazamiento
  X(6) = X(2): Y(6) = Y(5)
  X(7) = Xpos: Y(7) = Ypos - 4 * Desplazamiento
  X(8) = Xpos - Desplazamiento: Y(8) = Y(5)
  X(9) = Xpos - 3 * Desplazamiento: Y(9) = Y(5)
  X(10) = Xpos - 2 * Desplazamiento: Y(10) = Ypos
  X(11) = X(9): Y(11) = Y(2)
  X(12) = X(8): Y(12) = Y(2)
  Line (X(1), Y(1))-(X(5), Y(5)), Color
  Line -(X(9), Y(9)), Color
  Line -(X(1), Y(1)), Color
  Line (X(3), Y(3))-(X(7), Y(7)), Color
  Line -(X(11), Y(11)), Color
  Line -(X(3), Y(3)), Color
  Koch Xpos, Ypos, 2 * Desplazamiento
  For i = 1 To 12
    Koch X(i), Y(i), 3 * Desplazamiento
  Next i
End Sub

' DIBUJA LA CURVA C
Private Sub cmdCurva_Click()
  Dim X1!, Y1!, X2!, Y2!
  Cls
  Scale (-600, 600)-(600, -600)
  X1! = -250: Y1! = -200
  X2! = 250: Y2! = -200
  Curva X1!, Y1!, X2!, Y2!
End Sub

' SUBRUTINA PARA DIBUJAR LA CURVA C
Sub Curva(X1!, Y1!, X2!, Y2!)
  Dim DistanciaPuntos As Single, SiguienteX1 As Single, SiguienteY1 As Single
  Dim Color
  'Color = QBColor(Int(15 * Rnd))
  Color = vbWhite

  DistanciaPuntos = Distancia(X1!, Y1!, X2!, Y2!)
  If DistanciaPuntos < 3 Then
    Exit Sub
  End If
  SiguienteX1 = X1 + MovimientoX(X1!, Y1!, X2!, Y2!)
  SiguienteY1 = Y1 + MovimientoY(X1!, Y1!, X2!, Y2!)
  Line (X1!, Y1!)-(X2!, Y2!), BackColor
  Line (X1!, Y1!)-(SiguienteX1, SiguienteY1), Color
  Line -(X2!, Y2!), Color
  Curva X1!, Y1!, SiguienteX1, SiguienteY1
  Curva SiguienteX1, SiguienteY1, X2!, Y2!
End Sub

' FUNCION PARA DETERMINAR EL MOVIMIENTO DE LA CURVA C EN EJE X
Function MovimientoX(X1!, Y1!, X2!, Y2!) As Single
' Variables Locales
  Dim Angulo As Single, D As Single
  Dim DesplazamientoX As Single, DesplazamientoY As Single
  Dim R
  Angulo = Radianes(45)
  D = Distancia(X1!, Y1!, X2!, Y2!)
  R = (Sqr(2) / 2) * D
  DesplazamientoX = Cos(Angulo) * (X2! - X1!)
  DesplazamientoY = Sin(Angulo) * (Y2! - Y1!)
  MovimientoX = (R / D) * (DesplazamientoX - DesplazamientoY)
End Function

' FUNCION PARA DETERMINAR EL MOVIMIENTO DE LA CURVA C EN EJE Y
Function MovimientoY(X1!, Y1!, X2!, Y2!) As Single
' Variables Locales
  Dim Angulo As Single, D As Single
  Dim DesplazamientoX As Single, DesplazamientoY As Single
  Dim R
  Angulo = Radianes(45)
  D = Distancia(X1!, Y1!, X2!, Y2!)
  R = (Sqr(2) / 2) * D
  DesplazamientoX = Sin(Angulo) * (X2! - X1!)
  DesplazamientoY = Cos(Angulo) * (Y2! - Y1!)
  MovimientoY = (R / D) * (DesplazamientoX + DesplazamientoY)
End Function

' FUNCION PARA DETERMINAR LA DISTANCIA ENTRE DOS PUNTOS
Function Distancia(X1!, Y1!, X2!, Y2!) As Single
' Calcula la distancia entre dos puntos
  Dim X As Single, Y As Single
  X = (X2! - X1!) * (X2! - X1!)
  Y = (Y2! - Y1!) * (Y2! - Y1!)
  Distancia = Sqr(X + Y)
End Function

'FUNCION PARA CONVERTIR GRADOS EN RADIANES
Function Radianes(X!)
' Convierte los grados en radianes
  Radianes = X! * PI / 180
End Function

' DIBUJA LA CURVA DEL DRAGON
Private Sub cmdDragon_Click()
  Dim X1!, Y1!, X2!, Y2!
  Cls
  Scale (-900, 900)-(900, -900)
  X1! = -1000 / 2.2: Y1! = -800 / 10
  X2! = 1000 / 2.2: Y2! = -800 / 10
  CurvaDragon X1!, Y1!, X2!, Y2!
  CurvaDragon X2!, Y2!, X1!, Y1!
End Sub

' SUBRUTINA PARA DIBUJAR LA CURVA DEL DRAGON
Sub CurvaDragon(X1!, Y1!, X2!, Y2!)
  Dim DistanciaPuntos As Single, SiguienteX1 As Single, SiguienteY1 As Single

  Dim Color
  'Color = QBColor(Int(15 * Rnd))
  Color = vbWhite

  DistanciaPuntos = Distancia(X1!, Y1!, X2!, Y2!)
  If DistanciaPuntos < 5 Then
    Exit Sub
  End If
  SiguienteX1 = X1 + MovimientoDragonX(X1!, Y1!, X2!, Y2!)
  SiguienteY1 = Y1 + MovimientoDragonY(X1!, Y1!, X2!, Y2!)
  Line (X1!, Y1!)-(X2!, Y2!), BackColor
  Line (X1!, Y1!)-(SiguienteX1, SiguienteY1), Color
  Line -(X2!, Y2!), Color
  CurvaDragon X1!, Y1!, SiguienteX1, SiguienteY1
  CurvaDragon SiguienteX1, SiguienteY1, X2!, Y2!
End Sub

' FUNCION PARA DETERMINAR EL MOVIMIENTO DEL DRAGON EN EJE X
Function MovimientoDragonX(X1!, Y1!, X2!, Y2!) As Single
' Variables Locales
  Dim Angulo As Single, D As Single
  Dim DesplazamientoX As Single, DesplazamientoY As Single
  Dim R
  Static j As Integer
  j = j + 1
  Angulo = Radianes(45)
  If j Mod 2 = 0 Then Angulo = -Angulo: j = 0
  D = Distancia(X1!, Y1!, X2!, Y2!)
  R = (Sqr(2) / 2) * D
  DesplazamientoX = Cos(Angulo) * (X2! - X1!)
  DesplazamientoY = Sin(Angulo) * (Y2! - Y1!)
  MovimientoDragonX = (R / D) * (DesplazamientoX - DesplazamientoY)
End Function

' FUNCION PARA DETERMINAR EL MOVIMIENTO DEL DRAGON EN EJE Y
Function MovimientoDragonY(X1!, Y1!, X2!, Y2!) As Single
' Variables Locales
  Dim Angulo As Single, D As Single
  Dim DesplazamientoX As Single, DesplazamientoY As Single
  Dim R
  Static j As Integer
  j = j + 1
  Angulo = Radianes(45)
  If j Mod 2 = 0 Then Angulo = -Angulo: j = 0
  D = Distancia(X1!, Y1!, X2!, Y2!)
  R = (Sqr(2) / 2) * D
  DesplazamientoX = Sin(Angulo) * (X2! - X1!)
  DesplazamientoY = Cos(Angulo) * (Y2! - Y1!)
  MovimientoDragonY = (R / D) * (DesplazamientoX + DesplazamientoY)
End Function

' DIBUJA LOS CUADROS
Private Sub cmdCuadros_Click()
  Cls
  Scale (-2500, 2500)-(2500, -2500)
  Cuadro -1000, 1000, 2000
End Sub

' SUBRUTINA PARA DIBUJAR LOS CUADROS
Sub Cuadro(X, Y, Tamano)
  If Tamano < 40 Then Exit Sub
  Line (X, Y)-(X + Tamano, Y - Tamano), vbWhite, B
  Cuadro X - Tamano / 4, Y + Tamano / 4, Tamano / 2
  Cuadro X + Tamano - Tamano / 4, Y + Tamano / 4, Tamano / 2
  Cuadro X - Tamano / 4, Y - Tamano + Tamano / 4, Tamano / 2
  Cuadro X + Tamano - Tamano / 4, Y - Tamano + Tamano / 4, Tamano / 2
End Sub


' DIBUJA UN PAISAJE
Private Sub cmdPaisaje_Click()
  Dim X1!, Y1!, X2!, Y2!
  Cls
  Scale (-800, 800)-(800, -800)
  X1! = -330: Y1! = -130
  X2! = 330: Y2! = -130
  'CurvaPaisaje X1!, Y1!, X2!, Y2!
  CurvaPaisaje X2!, Y2!, 0, -Y2!
  CurvaPaisaje 0, -Y2!, X1!, Y1!
  CurvaPaisaje -X1!, Y1!, -X2!, Y2!
End Sub

' SUBRUTINA PARA DIBUJAR LA CURVA DEL Paisaje
Sub CurvaPaisaje(X1!, Y1!, X2!, Y2!)
  Dim DistanciaPuntos As Single, SiguienteX1 As Single, SiguienteY1 As Single

  Dim Color
  'Color = QBColor(Int(15 * Rnd))
  Color = vbWhite

  DistanciaPuntos = Distancia(X1!, Y1!, X2!, Y2!)
  If DistanciaPuntos < 5 Then
    Exit Sub
  End If
  SiguienteX1 = X1 + MovimientoPaisajeX(X1!, Y1!, X2!, Y2!)
  SiguienteY1 = Y1 + MovimientoPaisajeY(X1!, Y1!, X2!, Y2!)
  Line (X1!, Y1!)-(X2!, Y2!), BackColor
  Line (X1!, Y1!)-(SiguienteX1, SiguienteY1), Color
  Line -(X2!, Y2!), Color
  CurvaPaisaje X1!, Y1!, SiguienteX1, SiguienteY1
  CurvaPaisaje SiguienteX1, SiguienteY1, X2!, Y2!
End Sub

' FUNCION PARA DETERMINAR EL MOVIMIENTO DEL Paisaje EN EJE X
Function MovimientoPaisajeX(X1!, Y1!, X2!, Y2!) As Single
' Variables Locales
  Dim Angulo As Single, D As Single
  Dim DesplazamientoX As Single, DesplazamientoY As Single
  Dim R
  Angulo = Radianes(15 + 60 * Rnd)
  If Rnd(1) > 0.5 Then Angulo = -Angulo
  D = Distancia(X1!, Y1!, X2!, Y2!)
  R = (0.15 + (0.6 * Rnd(1))) * D
  DesplazamientoX = Cos(Angulo) * (X2! - X1!)
  DesplazamientoY = Sin(Angulo) * (Y2! - Y1!)
  MovimientoPaisajeX = (R / D) * (DesplazamientoX - DesplazamientoY)
End Function

' FUNCION PARA DETERMINAR EL MOVIMIENTO DEL Paisaje EN EJE Y
Function MovimientoPaisajeY(X1!, Y1!, X2!, Y2!) As Single
' Variables Locales
  Dim Angulo As Single, D As Single
  Dim DesplazamientoX As Single, DesplazamientoY As Single
  Dim R
  Angulo = Radianes(15 + 60 * Rnd)
  If Rnd(1) > 0.5 Then Angulo = -Angulo
  D = Distancia(X1!, Y1!, X2!, Y2!)
  R = (0.15 + (0.6 * Rnd(1))) * D
  DesplazamientoX = Sin(Angulo) * (X2! - X1!)
  DesplazamientoY = Cos(Angulo) * (Y2! - Y1!)
  MovimientoPaisajeY = (R / D) * (DesplazamientoX + DesplazamientoY)
End Function

