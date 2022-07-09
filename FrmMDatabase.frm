VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form FrmMDatabase 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Master Database"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   9000
      TabIndex        =   3
      Top             =   0
      Width           =   9030
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Base de Datos | Ultima Actualización: "
         BeginProperty Font 
            Name            =   "Ebrima"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   375
         Left            =   135
         TabIndex        =   4
         Top             =   0
         Width           =   8700
      End
   End
   Begin VB.CommandButton CmdImport 
      Caption         =   "&Importar / Actualizar base desde archivo CSV"
      Height          =   375
      Left            =   90
      TabIndex        =   2
      Top             =   7200
      Width           =   3615
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   7560
      TabIndex        =   1
      Top             =   7200
      Width           =   1365
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   6180
      Left            =   90
      TabIndex        =   0
      Top             =   585
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   10901
      _Version        =   393216
      Cols            =   5
      WordWrap        =   -1  'True
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   3
   End
   Begin MSComDlg.CommonDialog EstCMD 
      Left            =   4320
      Top             =   7020
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmMDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FUNCION PARA IMPORTAR UN CSV Y ACTUALIZAR O GENERAR LA BASE DE DATOS PRINCIPAL
'------------------------------------------------------------------------------------------------------

Private Sub ImportarCSV(Grid1 As MSFlexGrid, CSVEntrada As String, Optional Separador As String = ";")
  
  ' Carga en un FlexGrid un fichero en formato CSV
  Dim Fichero As Integer, Registro As String, Campos() As String, fila As Single, Columna As Single, Result As Boolean
  Dim CSVid As Long
  
  Grid1.Redraw = False 'grilla de carga
  'PrgBar1.Visible = True

  'preparaciones para lectura / escritura
  Fichero = FreeFile
  
  ' Archivo de entrada y salida
  Open CSVEntrada For Input As #Fichero
  ' Procesamos los datos de entrada hasta el final
  While Not EOF(Fichero)
    ' Leemos un Registro y lo separamos en Campos individuales
    Line Input #Fichero, Registro
    Campos = Split(Registro, Separador)
    ' Si es la primera Lectura (Fila=0) dimensionamos adecuadamente el Grid
    If fila = 0 Then
      Grid1.Clear                     'Lo borramos
      Grid1.FixedCols = 0             'Numero de Columnas fijas
      Grid1.FixedRows = 1             'Numero de Filas Fijas (Titulos)
      Grid1.Rows = 1                  '1 Fila
      Grid1.Cols = UBound(Campos) + 1
    End If
    ' Control de Fila a utilizar, se añade si es necesario
    If Grid1.Rows <= fila Then Grid1.Rows = fila + 1
    ' Situamos una a una las Columnas.
    For Columna = 0 To UBound(Campos)
        Grid1.TextMatrix(fila, Columna) = Trim(Campos(Columna))
    Next
    
    DoEvents
    
    'verificamos si es la fila 0 del archivo la que esta en posicion
    'ya que la fila 0 correspode a los encabezados de columna.
    If fila > 0 Then
        '/// seteamos los datos a guardar
        StockDatabase.Aid = CLng(Campos(0))
            CSVid = CLng(Campos(0))
        StockDatabase.CodeSec = Trim(Campos(1))
        StockDatabase.ArtName = Trim(Campos(2))
        StockDatabase.Exis = Trim(Campos(3))
        StockDatabase.Precio = Trim(Campos(4))
        '/// guardamos
        Result = GuardaInventario(StockDatabase, CSVid)
        '/// verificamos que este todo correcto
        If Result = False Then
            MsgBox "error"
            Close #Fichero
            Exit Sub
        End If
        Grid1.TextMatrix(fila, 0) = Trim(StockDatabase.Aid)
        Grid1.TextMatrix(fila, 1) = Trim(StockDatabase.CodeSec)
        Grid1.TextMatrix(fila, 2) = Trim(StockDatabase.ArtName)
        Grid1.TextMatrix(fila, 3) = Trim(StockDatabase.Exis)
        Grid1.TextMatrix(fila, 4) = Trim(StockDatabase.Precio)
    Else
        Grid1.TextMatrix(0, 0) = "SKU"
        Grid1.TextMatrix(0, 1) = "CODSEC"
        Grid1.TextMatrix(0, 2) = "ARTICULO"
        Grid1.TextMatrix(0, 3) = "EXIS"
        Grid1.TextMatrix(0, 4) = "PRECIO"
    End If

    ' Aumentamos número de Fila
    fila = fila + 1
    FrmSearching.Label1.Caption = "Procesando: " & CStr(fila) & " registros."
    
    'If PrgBar1.Value < Fila Then
        'PrgBar1.Value = Fila
    'End If
    
  Wend
  
  '// cerramos los ficheros CSV de entrada y salida
  Close #Fichero
  
  'algunas habilitaciones necesarias
  Grid1.Redraw = True
  'PrgBar1.Visible = False

End Sub

Private Sub CmdClose_Click()

Unload Me

End Sub

Private Sub CmdImport_Click()

Dim ConverTx As String

'abrir CSV origen
'On Error Resume Next
EstCMD.InitDir = App.Path & "\data\"
EstCMD.Filter = "Archivo CSV (*.csv)|*.csv|Archivos CSV"
EstCMD.DialogTitle = "Lib Lerma Stock - Importar archivo CSV"
EstCMD.CancelError = True
EstCMD.ShowOpen

If err.Number = 32755 Then Exit Sub

ConverTx = EstCMD.FileName

DoEvents

CmdClose.Enabled = False
CmdImport.Enabled = False

FrmSearching.Show , Me
FrmSearching.Label1.Caption = "Procesando..."

DoEvents
ImportarCSV Grid2, ConverTx

'ExportarCSV Grilla1, Grilla2, App.Path & "\data\database2.mdb", ConverTx, ConvertXX

CmdClose.Enabled = True
CmdImport.Enabled = True

Unload FrmSearching

MsgBox "Base de Datos actualizada correctamente!."

End Sub

Private Sub Form_Load()

Dim temporal As String

'// cargamos la ultima fecha de modificacion de la base de datos
temporal = App.Path & InvFilename
Label1.Caption = "Base de Datos Maestra / Ultima Actualizacion: " & GetFechaArch(temporal)


End Sub
