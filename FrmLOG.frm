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
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   11835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Abrir URL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5850
      TabIndex        =   5
      Top             =   3150
      Width           =   1290
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Definir URL de acceso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3930
      TabIndex        =   4
      Top             =   3150
      Width           =   1860
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9945
      TabIndex        =   1
      Top             =   3150
      Width           =   1770
   End
   Begin VB.ListBox List1 
      Height          =   2940
      ItemData        =   "FrmLOG.frx":0000
      Left            =   30
      List            =   "FrmLOG.frx":0002
      TabIndex        =   0
      Top             =   90
      Width           =   11670
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "---"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   2790
      TabIndex        =   3
      Top             =   3180
      Width           =   1065
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo de Producto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   3180
      Width           =   2565
   End
End
Attribute VB_Name = "FrmLOG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
 
Const SW_NORMAL = 1

Private Sub Command2_Click()

Dim X, CodProd

'Debug.Print List1.Text

'!! STOCK: 042254 -  Medidas A5 - TNDB: 2 >> IBER: 0
CodProd = Mid$(List1.Text, 11, 6)
'Debug.Print Mid$(List1.Text, 11, 6)

X = ShellExecute(Me.hwnd, "Open", "https://librerialerma2.mitiendanube.com/admin/products/?sort=created-descending&q=" & CodProd, &O0, &O0, SW_NORMAL)

End Sub

Private Sub Command3_Click()

Unload Me

End Sub

Private Sub List1_Click()

Label2.Caption = Mid$(List1.Text, 11, 6)

End Sub

Private Sub List1_DblClick()

Dim X, CodProd

'Debug.Print List1.Text

'!! STOCK: 042254 -  Medidas A5 - TNDB: 2 >> IBER: 0
CodProd = Mid$(List1.Text, 11, 6)
'Debug.Print Mid$(List1.Text, 11, 6)

X = ShellExecute(Me.hwnd, "Open", "https://librerialerma2.mitiendanube.com/admin/products/?sort=created-descending&q=" & CodProd, &O0, &O0, SW_NORMAL)

End Sub
