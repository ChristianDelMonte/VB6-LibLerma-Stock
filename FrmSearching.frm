VERSION 5.00
Begin VB.Form FrmSearching 
   BackColor       =   &H0080FF80&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " --- ESPERE POR FAVOR..."
   ClientHeight    =   480
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3795
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   480
   ScaleWidth      =   3795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BUSCANDO..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   3615
   End
End
Attribute VB_Name = "FrmSearching"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
