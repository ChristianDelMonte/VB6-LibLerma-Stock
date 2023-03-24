VERSION 5.00
Begin VB.Form FrmInit 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Prometheus - Libreria Lerma Stock"
   ClientHeight    =   7515
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   10695
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "FrmInit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   10695
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   10665
      TabIndex        =   0
      Top             =   7185
      Width           =   10695
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SISTEMA CARGADO Y LISTO PARA OPERAR"
         BeginProperty Font 
            Name            =   "Corbel"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   45
         TabIndex        =   1
         Top             =   45
         Width           =   4560
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      DrawStyle       =   5  'Transparent
      ForeColor       =   &H80000008&
      Height          =   2610
      Left            =   1530
      Picture         =   "FrmInit.frx":0442
      ScaleHeight     =   2610
      ScaleWidth      =   7500
      TabIndex        =   2
      Top             =   1710
      Width           =   7500
   End
   Begin VB.Menu sis 
      Caption         =   "&Sistema"
      Index           =   0
      Begin VB.Menu sis_vueltos 
         Caption         =   "Bajas"
         Index           =   1
         Begin VB.Menu sis_cvueltos 
            Caption         =   "Ingresos..."
            Index           =   1
            Shortcut        =   ^I
         End
         Begin VB.Menu sis_cconsultas 
            Caption         =   "Consultas..."
            Index           =   2
            Shortcut        =   ^C
         End
      End
      Begin VB.Menu sis_cplanillas 
         Caption         =   "Comparar Planillas..."
         Index           =   2
         Shortcut        =   ^P
      End
      Begin VB.Menu sis_sep1 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu sis_exit 
         Caption         =   "Salir"
         Index           =   4
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu act 
      Caption         =   "&Actualizaciones"
      Index           =   1
      Begin VB.Menu act_csvtienda 
         Caption         =   "CSV Tienda Nube..."
         Index           =   1
         Shortcut        =   ^T
      End
      Begin VB.Menu act_separador 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu act_masterbase 
         Caption         =   "Master Database..."
         Index           =   3
         Shortcut        =   ^M
      End
   End
End
Attribute VB_Name = "FrmInit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Centrar(control As Object)
  
    Dim contenedor As Object
  
    'Referencia al contenedor del control a centrar
    Set contenedor = control.Container
      
    With contenedor
        'Posición Top
        control.Top = ((.ScaleHeight - control.Height) / 2)
  
        'Posición Izquierda
        control.Left = ((.ScaleWidth - control.Width) / 2)
  
    End With
      
    'Eliminamos la referencia
    Set contenedor = Nothing
  
End Sub

Private Sub act_csvtienda_Click(Index As Integer)

FrmImport.Show , Me

End Sub

Private Sub act_masterbase_Click(Index As Integer)

FrmMDatabase.Show , Me

End Sub

Private Sub Form_Resize()

'Label1.Width = Picture1.Width - 100
Label2.Width = Picture2.Width - 100

Call Centrar(Picture3)

End Sub

Private Sub Picture2_Resize()

Label2.Width = Picture2.Width - 100

End Sub

Private Sub sis_cconsultas_Click(Index As Integer)

IngVCons.Show , Me

End Sub

Private Sub sis_cplanillas_Click(Index As Integer)

'frmMain.Show , Me

End Sub

Private Sub sis_cvueltos_Click(Index As Integer)

IngVuelto.Show , Me

End Sub

Private Sub sis_exit_Click(Index As Integer)

End

End Sub
