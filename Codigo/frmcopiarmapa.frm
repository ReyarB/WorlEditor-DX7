VERSION 5.00
Begin VB.Form frmcopiarmapa 
   BorderStyle     =   0  'None
   Caption         =   "Copiado de Zonas"
   ClientHeight    =   4215
   ClientLeft      =   90
   ClientTop       =   2445
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmcopiarmapa.frx":0000
   ScaleHeight     =   4215
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton iralmapa 
      Caption         =   "Ir al Mapa"
      BeginProperty Font 
         Name            =   "Morpheus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1080
      TabIndex        =   1
      Top             =   2640
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox Mapa 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   2160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image4 
      Height          =   4230
      Left            =   0
      Picture         =   "frmcopiarmapa.frx":3D274
      Top             =   0
      Width           =   705
   End
   Begin VB.Image Image3 
      Height          =   4230
      Left            =   3720
      Picture         =   "frmcopiarmapa.frx":43E76
      Top             =   0
      Width           =   705
   End
   Begin VB.Image Image2 
      Height          =   630
      Left            =   720
      Picture         =   "frmcopiarmapa.frx":4AB97
      Top             =   3600
      Width           =   2955
   End
   Begin VB.Image Image1 
      Height          =   630
      Left            =   720
      Picture         =   "frmcopiarmapa.frx":51913
      Top             =   0
      Width           =   2955
   End
End
Attribute VB_Name = "frmcopiarmapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************
'Author: ^[ReyarB]^ 29/12/2006
'Last modified: 01/01/2018
'*************************************************

Public copiarelmapa As Integer


Public Sub Image1_Click()
'COPIADO ZONA AZUL "ARRIBA"

On Error Resume Next

Dim X As Integer
Dim y As Integer

' ARRIBA
Mapa.Text = 0
y = 7
For X = (9 + 1) To (92 - 1)
        If MapData(X, y).TileExit.Map > 0 Then
            Mapa.Text = MapData(X, y).TileExit.Map
            Exit For
        End If
Next
    SeleccionIX = 1
    SeleccionFX = 100
    SeleccionIY = 8 'hoy 9
    SeleccionFY = 14 'hoy 14
   
    copiarelmapa = Mapa.Text
    
Call CopiarSeleccionMapa

If copiarelmapa <= 0 Then
        MsgBox "Este Mapa es el final del mundo. No encuentro traslados.", vbInformation
        modPaneles.VerFuncion 4, False
        
        frmMain.MemoriaAuxiliar.Visible = True
        frmMain.Mapa.Visible = True
        frmMain.iralmapa.Visible = True
        
Exit Sub
End If

Call modpegadomap.Image1

MapInfo.Changed = 1
UserPos.X = 12
UserPos.y = 94
End Sub

Public Sub Image2_Click()
'COPIADO ZONA VERDE "ABAJO"

On Error Resume Next

Dim X As Integer
Dim y As Integer

' ABAJO
Mapa.Text = 0
y = 94
For X = (9 + 1) To (92 - 1)
        If MapData(X, y).TileExit.Map > 0 Then
            Mapa.Text = MapData(X, y).TileExit.Map
            Exit For
        End If
Next

'Call modEdicion.Quitar_Translados_ReyarB
'Call modEdicion.Bloqueo_Todos(0)

    SeleccionIX = 1
    SeleccionFX = 100
    SeleccionIY = 87
    SeleccionFY = 93 'hoy
    
    copiarelmapa = Mapa.Text

        
Call CopiarSeleccionMapa

If copiarelmapa <= 0 Then
        MsgBox "Este Mapa es el final del mundo. No encuentro traslados.", vbInformation
        frmMain.Image1.Visible = False
        frmMain.Image2.Visible = False
        frmMain.Image3.Visible = False
        frmMain.Image4.Visible = False

        frmMain.MemoriaAuxiliar.Visible = True
        frmMain.Mapa.Visible = True
        frmMain.iralmapa.Visible = True
    
        frmMain.COPIAR_GRH(0).Visible = False
        frmMain.COPIAR_GRH(1).Visible = False
        frmMain.COPIAR_GRH(2).Visible = False
        frmMain.COPIAR_GRH(3).Visible = False
Exit Sub
End If

Call modpegadomap.Image2

frmMapas.Image1.Visible = False
frmMapas.Image2.Visible = True
frmMapas.Image3.Visible = False
frmMapas.Image4.Visible = False
UserPos.X = 12
UserPos.y = 10
End Sub

Public Sub Image3_Click()
'COPIADO ZONA AMARILLA "DERECHA"

On Error Resume Next

Dim X As Integer
Dim y As Integer
'DERECHA
Mapa.Text = 0
X = 92
For y = (7 + 1) To (94 - 1)
        If MapData(X, y).TileExit.Map > 0 Then
            Mapa.Text = MapData(X, y).TileExit.Map
            Exit For
        End If
Next
'Call modEdicion.Quitar_Translados_ReyarB
'Call modEdicion.Bloqueo_Todos(0)

    SeleccionIX = 83
    SeleccionFX = 91
    SeleccionIY = 1
    SeleccionFY = 100
    
    copiarelmapa = Mapa.Text

    Call CopiarSeleccionMapa
    
    If copiarelmapa <= 0 Then
        MsgBox "Este Mapa es el final del mundo. No encuentro traslados.", vbInformation
        frmMain.Image1.Visible = False
        frmMain.Image2.Visible = False
        frmMain.Image3.Visible = False
        frmMain.Image4.Visible = False
        
        frmMain.MemoriaAuxiliar.Visible = True
        frmMain.Mapa.Visible = True
        frmMain.iralmapa.Visible = True
        
        frmMain.COPIAR_GRH(0).Visible = False
        frmMain.COPIAR_GRH(1).Visible = False
        frmMain.COPIAR_GRH(2).Visible = False
        frmMain.COPIAR_GRH(3).Visible = False
        Exit Sub
    End If
    
    Call modpegadomap.Image3
frmMapas.Image1.Visible = False
frmMapas.Image2.Visible = False
frmMapas.Image3.Visible = True
frmMapas.Image4.Visible = False
UserPos.X = 12
UserPos.y = 15
End Sub

Public Sub Image4_Click()
'COPIADO ZONA ROJA "IZQUIERDA"

On Error Resume Next

Dim X As Integer
Dim y As Integer
'DERECHA

Mapa.Text = 0
X = 9
For y = (7 + 1) To (94 - 1)
        If MapData(X, y).TileExit.Map > 0 Then
            Mapa.Text = MapData(X, y).TileExit.Map
            Exit For
        End If
Next

'Call modEdicion.Quitar_Translados_ReyarB
'Call modEdicion.Bloqueo_Todos(0)

    SeleccionIX = 10 'HOY 10
    SeleccionFX = 18 'HOY 18
    SeleccionIY = 1
    SeleccionFY = 100
    
    copiarelmapa = Mapa.Text
    
Call CopiarSeleccionMapa

    If copiarelmapa <= 0 Then
        MsgBox "Este Mapa es el final del mundo. No encuentro traslados.", vbInformation
        frmMain.Image1.Visible = False
        frmMain.Image2.Visible = False
        frmMain.Image3.Visible = False
        frmMain.Image4.Visible = False
        
        frmMain.MemoriaAuxiliar.Visible = True
        frmMain.Mapa.Visible = True
        frmMain.iralmapa.Visible = True
        
        frmMain.COPIAR_GRH(0).Visible = False
        frmMain.COPIAR_GRH(1).Visible = False
        frmMain.COPIAR_GRH(2).Visible = False
        frmMain.COPIAR_GRH(3).Visible = False
        Exit Sub
    End If

Call modpegadomap.Image4

frmMapas.Image1.Visible = False
frmMapas.Image2.Visible = False
frmMapas.Image3.Visible = False
frmMapas.Image4.Visible = True
UserPos.X = 85
UserPos.y = 15
End Sub
Public Sub CopiarSeleccionMapa()
'*************************************************
'Author: ^[ReyarB]^ 29/12/2006
'Last modified: 01/01/2018
'*************************************************

On Error GoTo Fallo

    'podria usar copy mem , pero por las dudas no XD
    Dim X As Integer
    Dim y As Integer
    
            'MapData(9, 7).Blocked = 0
            'MapData(92, 7).Blocked = 0
            'MapData(9, 94).Blocked = 0
            'MapData(92, 94).Blocked = 0
    'Seleccionando = False
    SeleccionAncho = Abs(SeleccionIX - SeleccionFX) + 1
    SeleccionAlto = Abs(SeleccionIY - SeleccionFY) + 1
    ReDim SeleccionMap(SeleccionAncho, SeleccionAlto) As MapBlock
    For X = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1
                
                If MapData(X + SeleccionIX, y + SeleccionIY).NPCIndex > 0 Then
                    EraseChar MapData(X + SeleccionIX, y + SeleccionIY).CharIndex
                    MapData(X + SeleccionIX, y + SeleccionIY).NPCIndex = 0
                End If
                    'MapData(X + SeleccionIX, y + SeleccionIY).OBJInfo.objindex = 0
                    'MapData(X + SeleccionIX, y + SeleccionIY).OBJInfo.Amount = 0
                    'MapData(X + SeleccionIX, y + SeleccionIY).ObjGrh.GrhIndex = 0
                    ' Quitar Translados
                    MapData(X + SeleccionIX, y + SeleccionIY).TileExit.Map = 0
                    MapData(X + SeleccionIX, y + SeleccionIY).TileExit.X = 0
                    MapData(X + SeleccionIX, y + SeleccionIY).TileExit.y = 0
                 ' Quitar Triggers
                    MapData(X + SeleccionIX, y + SeleccionIY).Trigger = 0
            SeleccionMap(X, y) = MapData(X + SeleccionIX, y + SeleccionIY)
        Next
    Next
'frmcopiarmapa.Mapa.Visible = True
'frmcopiarmapa.iralmapa.Visible = True
'frmcopiarmapa.Image1.Visible = False
'frmcopiarmapa.Image2.Visible = False
'frmcopiarmapa.Image3.Visible = False
'frmcopiarmapa.Image4.Visible = False


'Public Sub CambioMapa()
    If Mapa.Text = 0 Then
        frmMapas.Cls
        DoEvents
        Unload Me
        MapInfo.Changed = 0
        Call modMapIO.NuevoMapa
        modMapIO.AbrirMapa frmMain.Dialog.FileName
        DoEvents
        EngineRun = True
    Exit Sub
    
    End If

'*************************************************
'Author: ^[ReyarB]^ 29/12/2006
'Last modified: 01/01/2018
'*************************************************

On Error GoTo Fallo
    ' Selecciones
    Seleccionando = False
    SeleccionIX = 0
    SeleccionIY = 0
    SeleccionFX = 0
    SeleccionFY = 0
    ' Traslados
    Dim tTrans As WorldPos
 
    If tTrans.Map = 0 Then
        If LenB(frmMain.Dialog.FileName) > 0 Then
        
        
            If FileExist(PATH_Save & NameMap_Save & Mapa & ".map", vbArchive) = True Then
                Call modMapIO.NuevoMapa
                frmMain.Dialog.FileName = PATH_Save & NameMap_Save & Mapa & ".map"
                modMapIO.AbrirMapa frmMain.Dialog.FileName
                
                If WalkMode = True Then
                    MoveCharbyPos UserCharIndex, UserPos.X, UserPos.y
                    CharList(UserCharIndex).Heading = SOUTH
                End If
                frmMain.mnuReAbrirMapa.Enabled = True
            End If
        End If
    End If
'frmcopiarmapa.Mapa.Visible = False
'frmcopiarmapa.iralmapa.Visible = False
'frmcopiarmapa.Image1.Visible = True
'frmcopiarmapa.Image2.Visible = True
'frmcopiarmapa.Image3.Visible = True
'frmcopiarmapa.Image4.Visible = True
'frmMapas.Show
DoEvents
Unload Me
    Exit Sub

Fallo:
    MsgBox "DobleClick::Error " & Err.Number & " - " & Err.Description

End Sub


Private Sub Mapa_Change()
If Mapa.Text = 0 Then
    iralmapa.Caption = "Salir"
    MapInfo.Changed = 0
    Else
    iralmapa.Caption = "Ir al Mapa"
End If
End Sub
