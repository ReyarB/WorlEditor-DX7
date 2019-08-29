VERSION 5.00
Begin VB.Form frmConvert 
   Caption         =   "Conversor Integer > Long / Long > CSM"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   6765
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2160
      TabIndex        =   4
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Convertir Integer > Long"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2040
      Width           =   2295
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Automatizar proceso"
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   5040
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Convertir Long > CSM"
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Numero del mapa:"
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Mapa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   12
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   ".map"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   11
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Instrucciones:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   $"frmConvert.frx":0000
      Height          =   1095
      Left            =   1200
      TabIndex        =   9
      Top             =   0
      Width           =   5415
   End
   Begin VB.Label Info 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   6375
   End
   Begin VB.Label Label6 
      Caption         =   "Hasta:"
      Height          =   255
      Left            =   3720
      TabIndex        =   7
      Top             =   1560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label7 
      Caption         =   "Mapa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   ".map"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "frmConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Automatico As Boolean

Private Sub Check1_Click()
    If Check1.value = False Then
        Label6.Visible = False
        Label7.Visible = False
        Label8.Visible = False
        Text2.Visible = False
        Automatico = False
    Else
        Label6.Visible = True
        Label7.Visible = True
        Label8.Visible = True
        Text2.Visible = True
        Automatico = True
    End If
End Sub

Private Sub Command1_Click()
Dim i As Integer
    If Automatico = False Then
        Call modMapIO.NuevoMapa
        Call MapaInteger_Cargar(App.Path & "\Conversor\Mapas Integer\Mapa" & Text1.Text & ".map")
        Call MapaV2_Guardar(App.Path & "\Conversor\Mapas Long\Mapa" & Text1.Text & ".map")
        
        Info.Caption = "Conversion realizada correctamente!"
    Else
        For i = Text1.Text To Text2.Text
            If FileExist(App.Path & "\Conversor\Mapas Integer\Mapa" & i & ".map", vbNormal) = True Then
                Call modMapIO.NuevoMapa
                Call MapaInteger_Cargar(App.Path & "\Conversor\Mapas Integer\Mapa" & i & ".map")
                Call MapaV2_Guardar(App.Path & "\Conversor\Mapas Long\Mapa" & i & ".map")
            
                Info.Caption = "Mapa" & i & " convertido correctamente!"
            End If
        Next i
    End If
End Sub

Private Sub Command2_Click()
Dim i As Integer
    If Automatico = False Then
        Call modMapIO.NuevoMapa
        Call MapaV2_Cargar(App.Path & "\Conversor\Mapas Long\Mapa" & Text1.Text & ".map")
        Call Save_CSM(App.Path & "\Conversor\Mapas CSM\Mapa" & Text1.Text & ".csm")
        
        Info.Caption = "Conversion realizada correctamente!"
    Else
        For i = Text1.Text To Text2.Text
            
            If FileExist(App.Path & "\Conversor\Mapas Long\Mapa" & i & ".map", vbNormal) = True Then
                Call modMapIO.NuevoMapa
                Call MapaV2_Cargar(App.Path & "\Conversor\Mapas Long\Mapa" & i & ".map")
                Call Save_CSM(App.Path & "\Conversor\Mapas CSM\Mapa" & i & ".csm")
            
                Info.Caption = "Mapa" & i & " convertido correctamente!"
            End If
        Next i
    End If
End Sub

