Attribute VB_Name = "modOptimizador"
Option Explicit


Public Sub Optimizardores()
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
        If (X < MinXBorder Or X > MaxXBorder Or y < MinYBorder Or y > MaxYBorder) Then
             'Quitar NPCs
            If MapData(X, y).NPCIndex > 0 Then
                EraseChar MapData(X, y).CharIndex
                MapData(X, y).NPCIndex = 0
            End If
            ' Quitar Translados
            MapData(X, y).TileExit.Map = 0
            MapData(X, y).TileExit.X = 0
            MapData(X, y).TileExit.y = 0
            ' Quitar Triggers
            MapData(X, y).Trigger = 0
        End If
        ' ** Quitar Translados y Triggers en Bloqueo
        If MapData(X, y).Blocked = 1 Then
            If MapData(X, y).TileExit.Map > 0 Then  ' Quita Translado Bloqueado
                MapData(X, y).TileExit.Map = 0
                MapData(X, y).TileExit.y = 0
                MapData(X, y).TileExit.X = 0
            ElseIf MapData(X, y).Trigger > 0 Then   ' Quita Trigger Bloqueado
                MapData(X, y).Trigger = 0
            End If
        End If
        ' ** Quitar Triggers en Translado
        If MapData(X, y).TileExit.Map > 0 Then
            If MapData(X, y).Trigger > 0 Then ' Quita Trigger en Translado
                MapData(X, y).Trigger = 0
            End If
        End If
        ' ** Mapea Arboles, Carteles, Foros y Yacimientos que no esten en la 3ra Capa
        If MapData(X, y).OBJInfo.objindex > 0 Then
            Select Case ObjData(MapData(X, y).OBJInfo.objindex).ObjType
                Case 4, 8, 10, 22 ' Arboles, Carteles, Foros, Yacimientos
                    If MapData(X, y).Graphic(3).GrhIndex <> MapData(X, y).ObjGrh.GrhIndex Then MapData(X, y).Graphic(3) = MapData(X, y).ObjGrh
                    If MapData(X, y).Blocked = 0 Then MapData(X, y).Blocked = 1 'ver hoy
            End Select
        End If

    Next X
Next y
    
'Set changed flag
MapInfo.Changed = 1
'ReyarB 2017
Call modEdicion.Bloquear_Bordes
'Fin ReyarB 2017

'************************************************
'Author: ^[GS]^
'Last modified: 22/09/06
'*************************************************
MapInfo.Changed = 1
DoEvents

End Sub



