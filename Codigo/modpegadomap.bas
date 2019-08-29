Attribute VB_Name = "modpegadomap"
Public Sub Image1()
'PEGAR ZONA AZUL ABAJO
    SobreX = 1
    SobreY = 94
Call modEdicion.Deshacer_Add("Pegar Selección")
Call pegadoReyarB
MapInfo.Changed = 1
DoEvents

End Sub
Public Sub Image2()
'PEGAR ZONA VERDE ARRIBA
    SobreX = 1
    SobreY = 1
Call modEdicion.Deshacer_Add("Pegar Selección")
Call pegadoReyarB
MapInfo.Changed = 1
DoEvents

End Sub
Public Sub Image3()
    SobreX = 1
    SobreY = 1
Call modEdicion.Deshacer_Add("Pegar Selección")
Call pegadoReyarB
MapInfo.Changed = 1
DoEvents

End Sub
Public Sub Image4()
If copiarelmapa <> 0 Then
Exit Sub
End If
    
    SobreX = 92
    SobreY = 1
Call modEdicion.Deshacer_Add("Pegar Selección")
Call pegadoReyarB
MapInfo.Changed = 1
DoEvents

End Sub

Private Sub pegadoReyarB()
On Error GoTo Fallo

Call frmUnionAdyacente.Show

    Static UltimoX As Integer
    Static UltimoY As Integer
    'If UltimoX = SobreX And UltimoY = SobreY Then Exit Sub
    UltimoX = SobreX
    UltimoY = SobreY
    Dim X As Integer
    Dim y As Integer
    DeSeleccionAncho = SeleccionAncho
    DeSeleccionAlto = SeleccionAlto
    DeSeleccionOX = SobreX
    DeSeleccionOY = SobreY
    ReDim DeSeleccionMap(DeSeleccionAncho, DeSeleccionAlto) As MapBlock
    
    For X = 0 To DeSeleccionAncho - 1
        For y = 0 To DeSeleccionAlto - 1
            DeSeleccionMap(X, y) = MapData(X + SobreX, y + SobreY)
        Next
    Next
    For X = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1
             MapData(X + SobreX, y + SobreY) = SeleccionMap(X, y)
        Next
    Next
    Seleccionando = False
    'Call DrawMiniMap(True)
    'frmOptimizador.Show
    Call modOptimizador.Optimizardores
    Call modEdicion.Bloquear_Bordes
    
            'MapData(9, 7).Blocked = 1
            'MapData(92, 7).Blocked = 1
            'MapData(9, 94).Blocked = 1
            'MapData(92, 94).Blocked = 1
    Exit Sub

Fallo:
    MsgBox "PegarSeleccion::Error " & Err.Number & " - " & Err.Description
    Call LogError("PegarSeleccion::Error " & Err.Number & " - " & Err.Description)
End Sub



