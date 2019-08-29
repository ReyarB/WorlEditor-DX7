VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form FrmListV 
   Caption         =   "Modificación de registros de BD"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8955
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FRMLISTV.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   8955
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CargarNPC 
      Caption         =   "Cargar NPCs"
      Height          =   255
      Left            =   4320
      TabIndex        =   3
      Top             =   240
      Width           =   3375
   End
   Begin MSComctlLib.ListView Lv1 
      Height          =   5535
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   9763
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CLose Me"
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Pulsa para editar registro y subregistros:"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   3390
   End
End
Attribute VB_Name = "FrmListV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CargarNPC_Click()
Call modIndices.CargarIndNPC
End Sub

Private Sub Command1_Click()


Unload Me

End Sub

'Private Sub Form_Load()
    
'Get the Path
'If Right(App.Path, 1) = "\" Then
'Path1 = App.Path
'Else
'Path1 = App.Path & "\"
'End If

'Call function to load the Data Base into the ListView
'LoadLv1


'End Sub


Private Sub Lv1_ItemClick(ByVal Item As MSComctlLib.ListItem)

'Put the Item to Edit in a Varible
Ace = Item
'Put the item and All SubItems in FramEdit textBoxes
FrmEdit.Text1 = Ace
FrmEdit.Text2 = Item.SubItems(1)
FrmEdit.Text3 = Item.SubItems(2)
FrmEdit.Text4 = Item.SubItems(3)
'FrmEdit.Text5 = Item.SubItems(4)
'FrmEdit.Text6 = Item.SubItems(5)
'Show FrmEdit
FrmEdit.Show
'Make FrmListV not Visible
FrmListV.Visible = False

End Sub



