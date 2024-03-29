VERSION 5.00
Begin VB.Form frmOptimizar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Optimizar Mapa"
   ClientHeight    =   4245
   ClientLeft      =   2085
   ClientTop       =   3390
   ClientWidth     =   3600
   Icon            =   "frmOptimizar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   3600
   Begin VB.CheckBox chkBloquearTrasladosAngulo 
      Caption         =   "Quitar Traslados y Bloqueos en angulos de los mapas"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   3375
   End
   Begin VB.CheckBox chkBloquearArbolesEtc 
      Caption         =   "Bloquear Arboles, Carteles, Foros y Yacimientos"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Value           =   1  'Checked
      Width           =   3375
   End
   Begin VB.CheckBox chkMapearArbolesEtc 
      Caption         =   "Mapear Arboles, Carteles, Foros y Yacimientos que no esten en la 3ra Capa"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   3375
   End
   Begin VB.CheckBox chkQuitarTodoBordes 
      Caption         =   "Quitar NPCs, Objetos y Translados en los Bordes Exteriores"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   3375
   End
   Begin VB.CheckBox chkQuitarTrigTrans 
      Caption         =   "Quitar Trigger's en Translados"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Value           =   1  'Checked
      Width           =   3375
   End
   Begin VB.CheckBox chkQuitarTrigBloq 
      Caption         =   "Quitar Trigger's Bloqueados"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Value           =   1  'Checked
      Width           =   3375
   End
   Begin VB.CheckBox chkQuitarTrans 
      Caption         =   "Quitar Translados Bloqueados"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Value           =   1  'Checked
      Width           =   3375
   End
   Begin WorldEditor.lvButtons_H cOptimizar 
      Default         =   -1  'True
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   3360
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1296
      Caption         =   "&Optimizar"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   12648384
   End
   Begin WorldEditor.lvButtons_H cCancelar 
      Height          =   735
      Left            =   1920
      TabIndex        =   6
      Top             =   3360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      Caption         =   "&Cancelar"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   1
      Value           =   0   'False
      cBack           =   -2147483633
   End
End
Attribute VB_Name = "frmOptimizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Optimizar()
'*************************************************
'Author: ^[GS]^
'Last modified: 16/10/06
'*************************************************
Dim y As Integer
Dim X As Integer

If Not MapaCargado Then
    Exit Sub
End If

' Quita Translados Bloqueados
' Quita Trigger's Bloqueados
' Quita Trigger's en Translados
' Quita NPCs, Objetos y Translados en los Bordes Exteriores
' Mapea Arboles, Carteles, Foros y Yacimientos que no esten en la 3ra Capa

modEdicion.Deshacer_Add "Aplicar Optimizacion del Mapa" ' Hago deshacer

For y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
        ' ** Quitar NPCs, Objetos y Translados en los Bordes Exteriores
        If (X < MinXBorder Or X > MaxXBorder Or y < MinYBorder Or y > MaxYBorder) And chkQuitarTodoBordes.value = 1 Then
             'Quitar NPCs
            If MapData(X, y).NPCIndex > 0 Then
                EraseChar MapData(X, y).CharIndex
                MapData(X, y).NPCIndex = 0
            End If
            ' Quitar Objetos
            MapData(X, y).OBJInfo.objindex = 0
            MapData(X, y).OBJInfo.Amount = 0
            MapData(X, y).ObjGrh.GrhIndex = 0
            ' Quitar Translados
            MapData(X, y).TileExit.Map = 0
            MapData(X, y).TileExit.X = 0
            MapData(X, y).TileExit.y = 0
            ' Quitar Triggers
            MapData(X, y).Trigger = 0
        End If
        ' ** Quitar Translados y Triggers en Bloqueo
        If MapData(X, y).Blocked = 1 Then
            If MapData(X, y).TileExit.Map > 0 And chkQuitarTrans.value = 1 Then ' Quita Translado Bloqueado
                MapData(X, y).TileExit.Map = 0
                MapData(X, y).TileExit.y = 0
                MapData(X, y).TileExit.X = 0
            ElseIf MapData(X, y).Trigger > 0 And chkQuitarTrigBloq.value = 1 Then ' Quita Trigger Bloqueado
                MapData(X, y).Trigger = 0
            End If
        End If
        ' ** Quitar Triggers en Translado
        If MapData(X, y).TileExit.Map > 0 And chkQuitarTrigTrans.value = 1 Then
            If MapData(X, y).Trigger > 0 Then ' Quita Trigger en Translado
                MapData(X, y).Trigger = 0
            End If
        End If
        ' ** Mapea Arboles, Carteles, Foros y Yacimientos que no esten en la 3ra Capa
        If MapData(X, y).OBJInfo.objindex > 0 And (chkMapearArbolesEtc.value = 1 Or chkBloquearArbolesEtc.value = 1) Then
            Select Case ObjData(MapData(X, y).OBJInfo.objindex).ObjType
                Case 4, 8, 10, 22 ' Arboles, Carteles, Foros, Yacimientos
                    If MapData(X, y).Graphic(3).GrhIndex <> MapData(X, y).ObjGrh.GrhIndex And chkMapearArbolesEtc.value = 1 Then MapData(X, y).Graphic(3) = MapData(X, y).ObjGrh
                    If chkBloquearArbolesEtc.value = 1 And MapData(X, y).Blocked = 0 Then MapData(X, y).Blocked = 1
            End Select
        End If
            MapData(9, 7).Blocked = 1
            MapData(92, 7).Blocked = 1
            MapData(9, 94).Blocked = 1
            MapData(92, 94).Blocked = 1
        ' ** Mapea Arboles, Carteles, Foros y Yacimientos que no esten en la 3ra Capa
    Next X
Next y
    If chkBloquearTrasladosAngulo.value = 1 Then
             ' Quitar Bloqueso en angulos
            MapData(9, 7).Blocked = 0
            MapData(92, 7).Blocked = 0
            MapData(9, 94).Blocked = 0
            MapData(92, 94).Blocked = 0

            ' Quitar Traslados en angulos
            MapData(9, 7).TileExit.Map = 0
            MapData(9, 7).TileExit.X = 0
            MapData(9, 7).TileExit.y = 0
            
            MapData(92, 7).TileExit.Map = 0
            MapData(92, 7).TileExit.X = 0
            MapData(92, 7).TileExit.y = 0
            
            MapData(9, 94).TileExit.Map = 0
            MapData(9, 94).TileExit.X = 0
            MapData(9, 94).TileExit.y = 0
            
            MapData(92, 94).TileExit.Map = 0
            MapData(92, 94).TileExit.X = 0
            MapData(92, 94).TileExit.y = 0
            
            MapInfo.Changed = 1
            End If

'Set changed flag
MapInfo.Changed = 1
'ReyarB 2017
Call modEdicion.Bloquear_Bordes
'Fin ReyarB 2017
End Sub

Private Sub cCancelar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 22/09/06
'*************************************************
Unload Me
End Sub

Private Sub cOptimizar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 22/09/06
'*************************************************
Call Optimizar
MapInfo.Changed = 1
DoEvents

Unload Me
End Sub

