VERSION 5.00
Begin VB.Form FrmLOG 
   BackColor       =   &H00808080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " --- VISOR DE EVENTOS"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11835
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   11835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   9945
      TabIndex        =   1
      Top             =   3150
      Width           =   1770
   End
   Begin VB.ListBox List1 
      Height          =   2985
      ItemData        =   "FrmLOG.frx":0000
      Left            =   45
      List            =   "FrmLOG.frx":0002
      TabIndex        =   0
      Top             =   90
      Width           =   11670
   End
End
Attribute VB_Name = "FrmLOG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()

Unload Me

End Sub
