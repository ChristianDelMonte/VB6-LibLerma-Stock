VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form IngVuelto 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  BAJAS - Ingreso / Consultas / Control"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6375
   Icon            =   "FrmVueltos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   6345
      TabIndex        =   27
      Top             =   0
      Width           =   6375
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "BAJAS | INGRESOS"
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   510
         Index           =   1
         Left            =   0
         TabIndex        =   28
         Top             =   0
         Width           =   6585
      End
   End
   Begin MSComCtl2.DTPicker Dtp1 
      Height          =   330
      Left            =   270
      TabIndex        =   24
      Top             =   1755
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      Format          =   112787457
      CurrentDate     =   43662
      MaxDate         =   47848
      MinDate         =   43465
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Cerrar"
      Height          =   420
      Left            =   5175
      TabIndex        =   23
      ToolTipText     =   "Cerrar y volver"
      Top             =   4950
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Consultas / Informes / Otros..."
      Height          =   420
      Left            =   135
      TabIndex        =   22
      ToolTipText     =   "Consultar los ingresos al sistema o generar informes"
      Top             =   4950
      Width           =   2715
   End
   Begin VB.Frame Frame2 
      Caption         =   "Carga de Articulos"
      Height          =   2445
      Left            =   135
      TabIndex        =   7
      Top             =   2340
      Width           =   6135
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000A&
         Caption         =   "<< &Guardar y continuar >>"
         Height          =   375
         Left            =   1980
         MaskColor       =   &H8000000F&
         TabIndex        =   14
         ToolTipText     =   "Guardar registro ingresado y continuar con la carga de uno nuevo"
         Top             =   1935
         Width           =   2175
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   360
         Left            =   135
         TabIndex        =   13
         Text            =   "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -"
         Top             =   1305
         Width           =   5865
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3825
         MaxLength       =   3
         TabIndex        =   12
         Text            =   "000"
         Top             =   360
         Width           =   1545
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   11
         Text            =   "000000"
         Top             =   360
         Width           =   1545
      End
      Begin VB.Label Label12 
         BackColor       =   &H00000000&
         Height          =   330
         Left            =   225
         TabIndex        =   18
         Top             =   1395
         Width           =   5820
      End
      Begin VB.Label Label11 
         BackColor       =   &H00000000&
         Height          =   375
         Left            =   3915
         TabIndex        =   17
         Top             =   450
         Width           =   1500
      End
      Begin VB.Label Label10 
         BackColor       =   &H00000000&
         Height          =   375
         Left            =   1125
         TabIndex        =   16
         Top             =   450
         Width           =   1545
      End
      Begin VB.Label Label9 
         BackColor       =   &H00000000&
         Height          =   330
         Left            =   2115
         TabIndex        =   15
         Top             =   2025
         Width           =   2085
      End
      Begin VB.Label Label6 
         Caption         =   "Cantidad:"
         Height          =   255
         Left            =   3015
         TabIndex        =   10
         Top             =   450
         Width           =   750
      End
      Begin VB.Label Label5 
         Caption         =   "Articulo:"
         Height          =   255
         Left            =   180
         TabIndex        =   9
         Top             =   1035
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Codigo:"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   450
         Width           =   645
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sucursal y Fecha"
      Height          =   1080
      Left            =   135
      TabIndex        =   0
      Top             =   1170
      Width           =   6135
      Begin VB.CommandButton Command4 
         Caption         =   "&Agregar..."
         Enabled         =   0   'False
         Height          =   330
         Left            =   2970
         TabIndex        =   25
         ToolTipText     =   "Agregar nuevas sucursales a la lista"
         Top             =   585
         Width           =   870
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   4275
         TabIndex        =   5
         Text            =   "1"
         Top             =   585
         Width           =   750
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1530
         TabIndex        =   4
         Text            =   "1"
         Top             =   585
         Width           =   1365
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         Height          =   285
         Left            =   3060
         TabIndex        =   26
         Top             =   675
         Width           =   825
      End
      Begin VB.Label Label15 
         BackColor       =   &H00000000&
         Height          =   285
         Left            =   4365
         TabIndex        =   21
         Top             =   675
         Width           =   735
      End
      Begin VB.Label Label14 
         BackColor       =   &H00000000&
         Height          =   285
         Left            =   1620
         TabIndex        =   20
         Top             =   675
         Width           =   1320
      End
      Begin VB.Label Label13 
         BackColor       =   &H00000000&
         Height          =   285
         Left            =   225
         TabIndex        =   19
         Top             =   675
         Width           =   1230
      End
      Begin VB.Label Label7 
         Caption         =   "1=Mañana 2=Tarde"
         ForeColor       =   &H000080FF&
         Height          =   420
         Left            =   5220
         TabIndex        =   6
         Top             =   540
         Width           =   825
      End
      Begin VB.Label Label3 
         Caption         =   "Turno:"
         Height          =   255
         Left            =   4290
         TabIndex        =   3
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Sucursal:"
         Height          =   255
         Left            =   1575
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha:"
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Label LblPrice 
      Caption         =   "$ 00000"
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
      Left            =   3105
      TabIndex        =   29
      Top             =   4995
      Width           =   1140
   End
End
Attribute VB_Name = "IngVuelto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    Text2.SetFocus
End If

End Sub


Private Sub Command1_Click()

'seteamos los datos a guardar
ControlVentas.Vdate = Dtp1.Value
ControlVentas.VSuc = Combo1.Text
ControlVentas.VTurno = CInt(Trim(Text2.Text))
ControlVentas.VICode = Trim(Text3.Text)
ControlVentas.VIDesc = Trim(Text5.Text)
ControlVentas.VICant = CInt(Trim(Text4.Text))
ControlVentas.VICosto = Trim(LblPrice.Caption)

If GuardarVentas(ControlVentas, -1) = False Then
    MsgBox "error"
Else
    'reiniciamos los textbox a cero para una nueva carga
    'Text1.Text = "00/00/0000"
    Text2.Text = "1"
    Text3.Text = "000000"
    Text4.Text = "000"
    Text5.Text = "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -"
    LblPrice.Caption = "$ 00000"
    'situamos el cursos en fecha para volver a cargar
    Dtp1.SetFocus
End If

End Sub

Private Sub Command2_Click()

IngVCons.Show , Me

End Sub

Private Sub Command3_Click()

Unload Me

End Sub

Private Sub Dtp1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

If KeyCode = 13 Then
    Combo1.SetFocus
End If


End Sub

Private Sub Form_Load()

'agregamos sucursales
Combo1.AddItem "1"
Combo1.AddItem "2"
Combo1.AddItem "3"
Combo1.AddItem "5"
Combo1.AddItem "6"
Combo1.AddItem "7"
Combo1.AddItem "8"
Combo1.AddItem "9"
Combo1.AddItem "10"
Combo1.AddItem "11"

End Sub

Private Sub Text2_Click()

Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)

End Sub

Private Sub Text2_GotFocus()

Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    Text3.SetFocus
End If

End Sub

Private Sub Text3_Click()

Text3.SelStart = 0
Text3.SelLength = Len(Text3.Text)

End Sub

Private Sub Text3_GotFocus()

Text3.SelStart = 0
Text3.SelLength = Len(Text3.Text)

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)

Dim Searcher As STKDatabase
Dim Cons As String, ConsCost As String
Dim Qbusco As Long

If KeyAscii = 13 Then
    DoEvents
    FrmSearching.Show , Me
    FrmSearching.Label1.Caption = "Buscando..."
    DoEvents
    '// consulta alternativa utilizando el MDB
    'Cons = BuscarArticuloMDB(App.Path & "\data\database2.mdb", Trim(Text3.Text))
    
    Qbusco = CLng(Trim(Text3.Text))
    Searcher = BuscarArticulo2(Qbusco)
    Cons = Trim(Searcher.ArtName)
    ConsCost = Trim(Searcher.Precio)
    If Cons <> "Error" Then
        Text5.Text = Cons
        LblPrice.Caption = ConsCost
        Text4.SetFocus
    Else
        Text5.Text = "NO ENCONTRADO"
        LblPrice.Caption = "$ 00000"
        Text3.SetFocus
    End If
    Unload FrmSearching
End If

End Sub

Private Sub Text4_Click()

Text4.SelStart = 0
Text4.SelLength = Len(Text4.Text)

End Sub

Private Sub Text4_GotFocus()

Text4.SelStart = 0
Text4.SelLength = Len(Text4.Text)

End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    Command1.SetFocus
End If

End Sub
