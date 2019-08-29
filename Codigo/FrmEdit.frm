VERSION 5.00
Begin VB.Form FrmEdit 
   Caption         =   "Editar NPC Auto Balance y manual"
   ClientHeight    =   7815
   ClientLeft      =   8985
   ClientTop       =   2265
   ClientWidth     =   10395
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmEdit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   10395
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Top             =   960
      Width           =   7815
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   4680
      Width           =   6015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar y salir"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   4200
      Width           =   6015
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   1680
      TabIndex        =   1
      Top             =   480
      Width           =   5295
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Nombre 
      Caption         =   "Nombre NPC"
      Height          =   195
      Left            =   1680
      TabIndex        =   8
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Oro"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Experiencia"
      Height          =   195
      Left            =   3360
      TabIndex        =   3
      Top             =   2400
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Npc Numero"
      Height          =   200
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1065
   End
End
Attribute VB_Name = "FrmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

