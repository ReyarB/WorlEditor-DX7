VERSION 5.00
Begin VB.Form frmNPCs 
   Caption         =   "NPCs"
   ClientHeight    =   8370
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7410
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "NPCs"
   ScaleHeight     =   8370
   ScaleWidth      =   7410
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cNumFunc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000012&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   330
      Index           =   2
      ItemData        =   "Form2.frx":0000
      Left            =   6240
      List            =   "Form2.frx":0002
      TabIndex        =   2
      Text            =   "1"
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton CargarNPC 
      Caption         =   "Importar lista de NPCs"
      Height          =   255
      Left            =   3360
      TabIndex        =   1
      Top             =   480
      Width           =   3735
   End
   Begin VB.ListBox lListadonpc 
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   6990
      Index           =   2
      ItemData        =   "Form2.frx":0004
      Left            =   240
      List            =   "Form2.frx":000B
      TabIndex        =   0
      Tag             =   "-1"
      Top             =   1320
      Width           =   6975
   End
   Begin VB.Image Image1 
      DragIcon        =   "Form2.frx":001C
      Height          =   855
      Left            =   240
      Picture         =   "Form2.frx":05A6
      Top             =   120
      Width           =   2250
   End
End
Attribute VB_Name = "frmNPCs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CargarNPC_Click()
Call modIndices.CargarIndNPC
End Sub

Private Sub cNumFunc_Change(Index As Integer)
NPC = Index
'Put the item and All SubItems in FramEdit textBoxes
FrmEdit.Text1 = Index
FrmEdit.Text2 = NpcData(Index).name
'Show FrmEdit
FrmEdit.Show
'Make FrmListV not Visible
FrmListV.Visible = False
End Sub

Private Sub lListadonpc_Click(Index As Integer)
NPC = Index
'Put the item and All SubItems in FramEdit textBoxes
FrmEdit.Text1 = Index
FrmEdit.Text2 = NpcData(Index).name
'Show FrmEdit
FrmEdit.Show
'Make FrmListV not Visible
FrmListV.Visible = False
End Sub
