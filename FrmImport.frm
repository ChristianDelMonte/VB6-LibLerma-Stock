VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form FrmImport 
   Caption         =   "Importar y Controlar Stock Web"
   ClientHeight    =   8385
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14670
   Icon            =   "FrmImport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8385
   ScaleWidth      =   14670
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   870
      Left            =   12060
      TabIndex        =   15
      Top             =   6885
      Width           =   2490
      Begin VB.CheckBox ChkInternal 
         Caption         =   "LLK"
         Height          =   240
         Left            =   315
         TabIndex        =   17
         Top             =   540
         Value           =   1  'Checked
         Width           =   870
      End
      Begin VB.CheckBox ChkExternal 
         Caption         =   "MDB"
         Enabled         =   0   'False
         Height          =   240
         Left            =   1440
         TabIndex        =   16
         Top             =   540
         Width           =   870
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Base de datos a Utilizar"
         Height          =   240
         Left            =   180
         TabIndex        =   18
         Top             =   225
         Width           =   2130
      End
   End
   Begin MSComDlg.CommonDialog EstCMDS 
      Left            =   11790
      Top             =   8055
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog EstCMD 
      Left            =   12015
      Top             =   7920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Height          =   1410
      Left            =   90
      TabIndex        =   9
      Top             =   6885
      Width           =   7125
      Begin VB.CheckBox Check1 
         Caption         =   "Mostrar LOG"
         Height          =   240
         Left            =   5715
         TabIndex        =   19
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1275
      End
      Begin VB.Label Label3 
         Caption         =   $"FrmImport.frx":0442
         Height          =   825
         Left            =   270
         TabIndex        =   10
         Top             =   225
         Width           =   6495
      End
   End
   Begin VB.Frame Frame1 
      Height          =   870
      Left            =   7425
      TabIndex        =   8
      Top             =   6885
      Width           =   4560
      Begin VB.CheckBox ChkBarras 
         Caption         =   "Codigo de Barras"
         Height          =   240
         Left            =   2655
         TabIndex        =   13
         Top             =   540
         Width           =   1635
      End
      Begin VB.CheckBox ChkStock 
         Caption         =   "Stock"
         Height          =   240
         Left            =   1530
         TabIndex        =   12
         Top             =   540
         Width           =   870
      End
      Begin VB.CheckBox ChkPrecio 
         Caption         =   "Precio"
         Enabled         =   0   'False
         Height          =   240
         Left            =   405
         TabIndex        =   11
         Top             =   540
         Value           =   1  'Checked
         Width           =   870
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Al exportar archivo CSV actualizar los siguientes campos:"
         Height          =   240
         Left            =   180
         TabIndex        =   14
         Top             =   225
         Width           =   4200
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   12735
      TabIndex        =   6
      Top             =   7920
      Width           =   1770
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Actualizar y Exportar..."
      Height          =   375
      Left            =   9405
      TabIndex        =   5
      Top             =   7920
      Width           =   2265
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Importar CSV..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   7425
      TabIndex        =   1
      Top             =   7920
      Width           =   1770
   End
   Begin MSFlexGridLib.MSFlexGrid Grilla1 
      Height          =   6180
      Left            =   7425
      TabIndex        =   0
      Top             =   585
      Width           =   7170
      _ExtentX        =   12647
      _ExtentY        =   10901
      _Version        =   393216
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   3
   End
   Begin MSFlexGridLib.MSFlexGrid Grilla2 
      Height          =   6180
      Left            =   90
      TabIndex        =   2
      Top             =   585
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   10901
      _Version        =   393216
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   3
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Procesando archivo CSV. Esta operacion puede tardar unos minutos! espere por favor..."
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   90
      TabIndex        =   7
      Top             =   90
      Visible         =   0   'False
      Width           =   14505
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      Caption         =   "Productos TIENDA NUBE"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   7440
      TabIndex        =   4
      Top             =   360
      Width           =   7170
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      Caption         =   "Productos IBERICO"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   90
      TabIndex        =   3
      Top             =   360
      Width           =   7125
   End
End
Attribute VB_Name = "FrmImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ShowLOG As Boolean

Private Sub ExportarCSVfromDB(Grid1 As MSFlexGrid, Grid2 As MSFlexGrid, CSVEntrada As String, CSVSalida As String, Optional Separador As String = ";")

    Dim s As String
    Dim n As Long
    Dim Lid As Long, Datos As STKDatabase
    Dim Mdbprecio As String, Csvprecio As String, Csvprecioprom As String
    Dim mdbnum As Double, csvnum As Double, CBarritas As String, CStock As Long, TStock As Long
    Dim msgito As String
    Dim mdbcalc As Double, mdbnuewnum As Double, PrecioProm As String
   
    DoEvents
    
  ' Carga en un FlexGrid un fichero en formato CSV
  Dim Fichero As Integer, Registro As String, Campos() As String, fila As Single, Columna As Single

  Grid1.Redraw = False 'grilla del csv
  Grid2.Redraw = False 'grilla del db
  
  Command1.Enabled = False
  Command2.Enabled = False
  Command3.Enabled = False
  Label4.Visible = True
  'prgb1.Visible = True

'//chequeamos si debemos mostrar el LOG o no
    If ShowLOG = True Then
        FrmLOG.Show , Me
        FrmLOG2.Show , Me
    End If
    
  'preparaciones para lectura / escritura
  Fichero = FreeFile
  
  ' Archivo de entrada y salida
  Open CSVEntrada For Input As #Fichero
  Open CSVSalida For Output As #33
    
  ' Procesamos los datos de entrada hasta el final
  Do While Not EOF(Fichero)
    ' Leemos un Registro y lo separamos en Campos individuales
    Line Input #Fichero, Registro
    Campos = Split(Registro, Separador)
    ' Si es la primera Lectura (Fila=0) dimensionamos adecuadamente el Grid
    If fila = 0 Then
      Grid1.Clear: Grid2.Clear                     'Lo borramos
      Grid1.FixedCols = 0: Grid2.FixedCols = 0             'Numero de Columnas fijas
      Grid1.FixedRows = 1: Grid2.FixedRows = 1              'Numero de Filas Fijas (Titulos)
      Grid1.Rows = 1: Grid2.Rows = 1                  '1 Fila
      Grid1.Cols = UBound(Campos) + 1
      Grid2.Cols = 5
    End If
    
    ' Control de Fila a utilizar, se añade si es necesario
    If Grid1.Rows <= fila Then Grid1.Rows = fila + 1: If Grid2.Rows <= fila Then Grid2.Rows = fila + 1
    ' Situamos una a una las Columnas.
    For Columna = 0 To UBound(Campos)
        Grid1.TextMatrix(fila, Columna) = Campos(Columna)
    Next
    
    DoEvents
    
    'verificamos si es la fila 0 del archivo la que esta en posicion
    'ya que la fila 0 correspode a los encabezados de columna.
    If fila > 0 Then
        'extraemos y verificamos que el articulo no tenga precio promocional
        Csvprecioprom = Campos(10)
        Csvprecioprom = Trim(Csvprecioprom)
        
        'extraemos el SKU para buscar en la database el articulo
        CodigoArt = Campos(16)      'campo 16 = SKU
        CodigoArt = Trim(CodigoArt)
        
        'verificamos el correcto formato de los codigos SKU
        If Len(CodigoArt) = 5 Then
            CodigoArt = "0" & CodigoArt
        ElseIf Len(CodigoArt) < 5 Then
            'pasar por alto el articulo
            msgito = "ATN! - SKU: " & Trim(Campos(1)) & " " & Trim(Campos(3)) & " " & Trim(Campos(4)) & " - SKIP!"
            If ShowLOG = True Then FrmLOG.List1.AddItem msgito
            GoTo Salto
        ElseIf Len(CodigoArt) > 6 Then
            'pasar por alto el articulo
            msgito = "ATN! - SKU: " & Trim(Campos(1)) & " " & Trim(Campos(3)) & " " & Trim(Campos(4)) & " - SKIP!"
            If ShowLOG = True Then FrmLOG.List1.AddItem msgito
            GoTo Salto
        End If

        'buscamos en la base de datos el SKU del articulo correspondiente al csv
        Datos = BuscarArticulo2(CLng(CodigoArt))
        
        'MsgBox tRs.Fields("codart") & " - $" & tRs.Fields("nombre") & " - $" & tRs.Fields("precio") & " - " & tRs.Fields("exis")
        Grid2.TextMatrix(fila, 0) = Trim(Datos.Aid)
        Grid2.TextMatrix(fila, 1) = Trim(Datos.ArtName)
        
        '// mdbnum = precio real del producto
        'On Error Resume Next
        
        'Debug.Print Datos.Precio
        mdbnum = CDbl(Trim(Datos.Precio))
        Mdbprecio = Format(mdbnum, "0.00")
        
        Grid2.TextMatrix(fila, 2) = Mdbprecio
        Grid2.TextMatrix(fila, 3) = Trim(Datos.Exis)
        
        '// mdbcalc = calculo del porcetnaje de precio promocional
        If mdbnum < 500 Then
            mdbcalc = (20 / 100) * mdbnum
        Else
            If mdbnum > 500 And mdbnum < 5000 Then
                mdbcalc = (15 / 100) * mdbnum
            Else
                If mdbnum > 5000 And mdbnum < 10000 Then
                    mdbcalc = (10 / 100) * mdbnum
                Else
                    If mdbnum > 10000 Then
                        mdbcalc = (5 / 100) * mdbnum
                    End If
                End If
            End If
        End If
        
        '//mdbnewnum = suma del % del precio + precio
        mdbnewnum = mdbnum + mdbcalc
        PrecioProm = Format(mdbnewnum, "0.00")
        
        'verificamos si hay que actualizar el codigo de barras del articulo
        If ChkBarras.Value = 1 Then
            CBarritas = CStr(Trim(Datos.CodeSec))
        Else
            CBarritas = Campos(17)
        End If
        
        'verificamos si hay que actualizar el stock del articulo
        If ChkStock.Value = 1 Then
            'deshabilitamos por el momento la actualizacion de stock
            'CStock = Campos(15)
        Else
            If Trim(Campos(15)) <> "" Then
                TStock = CLng(Trim(Campos(15)))     'stock tienda nube
                CStock = CLng(Trim(Datos.Exis))     'stock base de datos iberico
                'verificamos por incoherencias en el stock de ambos sistemas
                If TStock = 0 And TStock < CStock Then
                    If CStock > 2 Then
                        msgito = "!! STOCK: " & CodigoArt & " - " & Trim(Campos(1)) & " " & Trim(Campos(3)) & " " & Trim(Campos(4)) & " - TNDB: " & TStock & " << IBER: " & CStock
                        If ShowLOG = True Then FrmLOG2.List1.AddItem msgito
                    End If
                Else
                    If TStock > CStock And CStock <= 0 And TStock <> 0 Then
                        msgito = "!! STOCK: " & CodigoArt & " - " & Trim(Campos(1)) & " " & Trim(Campos(3)) & " " & Trim(Campos(4)) & " - TNDB: " & TStock & " >> IBER: " & CStock
                        If ShowLOG = True Then FrmLOG.List1.AddItem msgito
                    End If
                End If
            End If
        End If
                
        '/// exportamos los datos unificados entre el csv de origen y la database actualizada
        If Csvprecioprom = "" Then
            Write #33, Campos(0) & ";" & Campos(1) & ";" & Campos(2) & ";" & Campos(3) & _
            ";" & Campos(4) & ";" & Campos(5) & ";" & Campos(6) & ";" & Campos(7) & ";" & Campos(8) & _
            ";" & Mdbprecio & ";" & Campos(10) & ";" & Campos(11) & ";" & Campos(12) & ";" & Campos(13) & _
            ";" & Campos(14) & ";" & Campos(15) & ";" & CodigoArt & ";" & CBarritas & _
            ";" & Campos(18) & ";" & Campos(19) & ";" & Campos(20) & ";" & Campos(21) & ";" & Campos(22) & _
            ";" & Campos(23) & ";" & Campos(24)
        Else
            Write #33, Campos(0) & ";" & Campos(1) & ";" & Campos(2) & ";" & Campos(3) & _
            ";" & Campos(4) & ";" & Campos(5) & ";" & Campos(6) & ";" & Campos(7) & ";" & Campos(8) & _
            ";" & PrecioProm & ";" & Mdbprecio & ";" & Campos(11) & ";" & Campos(12) & ";" & Campos(13) & _
            ";" & Campos(14) & ";" & Campos(15) & ";" & CodigoArt & ";" & CBarritas & _
            ";" & Campos(18) & ";" & Campos(19) & ";" & Campos(20) & ";" & Campos(21) & ";" & Campos(22) & _
            ";" & Campos(23) & ";" & Campos(24)
        End If
    Else
        Grid2.TextMatrix(0, 0) = "codart"
        Grid2.TextMatrix(0, 1) = "nombre"
        Grid2.TextMatrix(0, 2) = "precio"
        Grid2.TextMatrix(0, 3) = "exis"
        'si es la fila 0 y estamos en los ancabezados de columna debemos exportar los mismos
        'para mantener el formato del archivo de origen.
        
Salto:
        Write #33, Campos(0) & ";" & Campos(1) & ";" & Campos(2) & ";" & Campos(3) _
        & ";" & Campos(4) & ";" & Campos(5) & ";" & Campos(6) & ";" & Campos(7) & ";" & Campos(8) _
        & ";" & Campos(9) & ";" & Campos(10) & ";" & Campos(11) & ";" & Campos(12) & ";" & Campos(13) _
        & ";" & Campos(14) & ";" & Campos(15) & ";" & Campos(16) & ";" & Campos(17) & ";" & Campos(18) _
        & ";" & Campos(19) & ";" & Campos(20) & ";" & Campos(21) & ";" & Campos(22) & ";" & Campos(23) _
        & ";" & Campos(24)
    End If

Salto2:
    ' Aumentamos número de Fila
    fila = fila + 1
    FrmSearching.Label1.Caption = "Procesando: " & fila
    'If prgb1.Value < Fila Then
    '    prgb1.Value = Fila
    'End If
  Loop
  
  '// cerramos los ficheros CSV de entrada y salida
  Close #Fichero
  Close #33
  
  'algunas habilitaciones necesarias
  Grid1.Redraw = True: Grid2.Redraw = True
  'prgb1.Visible = False
  Command1.Enabled = True
  Command2.Enabled = True
  Command3.Enabled = True
  Label4.Visible = False
  
End Sub

'FUNCION PARA EXPORTAR UN CSV
'CARGAMOS LOS DATOS DE LA PRIMARY DATABASE Y CARGAMOS EL CSV PRIMARIO Y LUEGO UNIFICAMOS LOS DATOS
'Y EXPORTAMOS EN UN NUEVO CSV.
'------------------------------------------------------------------------------------------------------
Private Sub ExportarCSVfromMDB(Grid1 As MSFlexGrid, Grid2 As MSFlexGrid, PrimaryMDB As String, CSVEntrada As String, CSVSalida As String, Optional Separador As String = ";")

    Dim s As String
    Dim tRs As ADODB.Recordset
    Dim n As Long
    Dim Mdbprecio As String, Csvprecio As String, Csvprecioprom As String
    Dim mdbnum As Double, csvnum As Double, CBarritas As String, CStock As Integer

    Set cnn = New ADODB.Connection
    '
    Provider = "Microsoft.Jet.OLEDB.4.0"
    cnn.Open "Provider=" & Provider & "; Data Source=" & PrimaryMDB
    '
    nombretabla = "VALORSTK"

    Set tRs = New ADODB.Recordset
    
    DoEvents
    
  ' Carga en un FlexGrid un fichero en formato CSV
  Dim Fichero As Integer, Registro As String, Campos() As String, fila As Single, Columna As Single

  Grid1.Redraw = False 'grilla del csv
  Grid2.Redraw = False 'grilla del mdb
  
  Command1.Enabled = False
  Command2.Enabled = False
  Command3.Enabled = False
  Label4.Visible = True
  'prgb1.Visible = True

  'preparaciones para lectura / escritura
  Fichero = FreeFile
  
  ' Archivo de entrada y salida
  Open CSVEntrada For Input As #Fichero
  Open CSVSalida For Output As #33
    
  ' Procesamos los datos de entrada hasta el final
  While Not EOF(Fichero)
    ' Leemos un Registro y lo separamos en Campos individuales
    Line Input #Fichero, Registro
    Campos = Split(Registro, Separador)
    ' Si es la primera Lectura (Fila=0) dimensionamos adecuadamente el Grid
    If fila = 0 Then
      Grid1.Clear: Grid2.Clear                     'Lo borramos
      Grid1.FixedCols = 0: Grid2.FixedCols = 0             'Numero de Columnas fijas
      Grid1.FixedRows = 1: Grid2.FixedRows = 1              'Numero de Filas Fijas (Titulos)
      Grid1.Rows = 1: Grid2.Rows = 1                  '1 Fila
      Grid1.Cols = UBound(Campos) + 1
      Grid2.Cols = 5
    End If
    ' Control de Fila a utilizar, se añade si es necesario
    If Grid1.Rows <= fila Then Grid1.Rows = fila + 1: If Grid2.Rows <= fila Then Grid2.Rows = fila + 1
    ' Situamos una a una las Columnas.
    For Columna = 0 To UBound(Campos)
        Grid1.TextMatrix(fila, Columna) = Campos(Columna)
    Next
    DoEvents
    'verificamos si es la fila 0 del archivo la que esta en posicion
    'ya que la fila 0 correspode a los encabezados de columna.
    If fila > 0 Then
        'extraemos y verificamos que el articulo no tenga precio promocional
        Csvprecioprom = Campos(10)
        Csvprecioprom = Trim(Csvprecioprom)
        'extraemos el SKU para buscar en la database el articulo
        CodigoArt = Campos(16)      'campo 16 = SKU
        CodigoArt = Trim(CodigoArt)
        'verificamos el correcto formato de los codigos SKU
        If Len(CodigoArt) = 5 Then
            CodigoArt = "0" & CodigoArt
        ElseIf Len(CodigoArt) < 5 Then
            'pasar por alto el articulo
        ElseIf Len(CodigoArt) > 6 Then
            'pasar por alto el articulo
        End If
        'buscamos en la base de datos el SKU del articulo correspondiente al csv
        s = "SELECT * FROM VALORSTK WHERE codart LIKE ""%" & CodigoArt & "%"";"
        tRs.Open s, cnn, adOpenDynamic, adLockOptimistic 'adLockOptimistic
        With tRs
            If (.EOF = True) And (.BOF = True) Then
                ' Si no hay datos...
                MsgBox "NADA ups!"
            Else
                n = 0
                Do While Not .EOF
                    n = n + 1
                    'MsgBox tRs.Fields("codart") & " - $" & tRs.Fields("nombre") & " - $" & tRs.Fields("precio") & " - " & tRs.Fields("exis")
                    Grid2.TextMatrix(fila, 0) = tRs.Fields("codart")
                    Grid2.TextMatrix(fila, 1) = tRs.Fields("nombre")
                    mdbnum = CDbl(tRs.Fields("precio"))
                    Mdbprecio = Format(mdbnum, "0.00")
                    Grid2.TextMatrix(fila, 2) = Mdbprecio
                    Grid2.TextMatrix(fila, 3) = tRs.Fields("exis")
                    'verificamos si hay que actualizar el codigo de barras del articulo
                    If ChkBarras.Value = 1 Then
                        CBarritas = CStr(tRs.Fields("codsec"))
                    Else
                        CBarritas = Campos(17)
                    End If
                    'verificamos si hay que actualizar el stock del articulo
                    'If ChkStock.Value = 1 Then
                        'CStock = CInt(tRs.Fields("exis"))
                        'deshabilitamos por el momento la actualizacion de stock
                    '    CStock = Campos(15)
                    'Else
                    '    CStock = Campos(15)
                    'End If
                    .MoveNext
                Loop
                '/// exportamos los datos unificados entre el csv de origen y la database actualizada
                If Csvprecioprom = "" Then
                    Write #33, Campos(0) & ";" & Campos(1) & ";" & Campos(2) & ";" & Campos(3) & _
                    ";" & Campos(4) & ";" & Campos(5) & ";" & Campos(6) & ";" & Campos(7) & ";" & Campos(8) & _
                    ";" & Mdbprecio & ";" & Campos(10) & ";" & Campos(11) & ";" & Campos(12) & ";" & Campos(13) & _
                    ";" & Campos(14) & ";" & Campos(15) & ";" & Campos(16) & ";" & CBarritas & _
                    ";" & Campos(18) & ";" & Campos(19) & ";" & Campos(20) & ";" & Campos(21) & ";" & Campos(22) & _
                    ";" & Campos(23) & ";" & Campos(24)
                Else
                    Write #33, Campos(0) & ";" & Campos(1) & ";" & Campos(2) & ";" & Campos(3) & _
                    ";" & Campos(4) & ";" & Campos(5) & ";" & Campos(6) & ";" & Campos(7) & ";" & Campos(8) & _
                    ";" & Campos(9) & ";" & Mdbprecio & ";" & Campos(11) & ";" & Campos(12) & ";" & Campos(13) & _
                    ";" & Campos(14) & ";" & Campos(15) & ";" & Campos(16) & ";" & CBarritas & _
                    ";" & Campos(18) & ";" & Campos(19) & ";" & Campos(20) & ";" & Campos(21) & ";" & Campos(22) & _
                    ";" & Campos(23) & ";" & Campos(24)
                End If
            End If
        End With
        tRs.Close
    Else
        Grid2.TextMatrix(0, 0) = "codart"
        Grid2.TextMatrix(0, 1) = "nombre"
        Grid2.TextMatrix(0, 2) = "precio"
        Grid2.TextMatrix(0, 3) = "exis"
        'si es la fila 0 y estamos en los ancabezados de columna debemos exportar los mismos
        'para mantener el formato del archivo de origen.
        Write #33, Campos(0) & ";" & Campos(1) & ";" & Campos(2) & ";" & Campos(3) _
        & ";" & Campos(4) & ";" & Campos(5) & ";" & Campos(6) & ";" & Campos(7) & ";" & Campos(8) _
        & ";" & Campos(9) & ";" & Campos(10) & ";" & Campos(11) & ";" & Campos(12) & ";" & Campos(13) _
        & ";" & Campos(14) & ";" & Campos(15) & ";" & Campos(16) & ";" & Campos(17) & ";" & Campos(18) _
        & ";" & Campos(19) & ";" & Campos(20) & ";" & Campos(21) & ";" & Campos(22) & ";" & Campos(23) _
        & ";" & Campos(24)
    End If

    ' Aumentamos número de Fila
    fila = fila + 1
    FrmSearching.Label1.Caption = "Procesando: " & fila
    'If prgb1.Value < Fila Then
    '    prgb1.Value = Fila
    'End If
    
  Wend
  
  '// cerramos los ficheros CSV de entrada y salida
  Close #Fichero
  Close #33
  
  'algunas habilitaciones necesarias
  Grid1.Redraw = True: Grid2.Redraw = True
  'prgb1.Visible = False
  Command1.Enabled = True
  Command2.Enabled = True
  Command3.Enabled = True
  Label4.Visible = False
  
  '// cerramos la coneccion con la Database Primaria.
  'tRs.Close
  cnn.Close

End Sub

'CAGAR O IMPORTAR CSV
'CARGAMOS LOS DATOS DE LA PRIMARY DATABASE Y CARGAMOS EL CSV PRIMARIO Y MOSTRAMOS LOS DATOS EN LAS GRILLAS
'CORRESPONDIENTES PARA SU ANALISIS.
'------------------------------------------------------------------------------------------------------
Private Sub CargarCSV(Grid1 As MSFlexGrid, Grid2 As MSFlexGrid, PrimaryMDB As String, FicheroCSV As String, Optional Separador As String = ";")

    Dim s As String
    Dim tRs As ADODB.Recordset
    Dim n As Long
    Dim Mdbprecio As String
    Dim mdbnum As Double
    Dim Csvprecio As String
    Dim csvnum As Double
    Dim Canal%
    
    Set cnn = New ADODB.Connection
    '
    Provider = "Microsoft.Jet.OLEDB.4.0"
    cnn.Open "Provider=" & Provider & "; Data Source=" & PrimaryMDB
    '
    nombretabla = "VALORSTK"

    Set tRs = New ADODB.Recordset
    
    DoEvents
    
  ' Carga en un FlexGrid un fichero en formato CSV
  Dim Fichero As Integer, Registro As String, Campos() As String, fila As Single, Columna As Single
  
  Grid1.Redraw = False 'grilla del csv
  Grid2.Redraw = False 'grilla del mdb
  
  Command1.Enabled = False
  Command2.Enabled = False
  Command3.Enabled = False
  Label4.Visible = True
  'prgb1.Visible = True

  ' Abrimos el fichero de Texto
  Fichero = FreeFile
  Open FicheroCSV For Input As #Fichero
  ' Lo procesamos hasta el final
  While Not EOF(Fichero)
    ' Leemos un Registro y lo separamos en Campos individuales
    Line Input #Fichero, Registro
    Campos = Split(Registro, Separador)
    ' Si es la primera Lectura (Fila=0) dimensionamos adecuadamente el Grid
    If fila = 0 Then
      Grid1.Clear: Grid2.Clear                     'Lo borramos
      Grid1.FixedCols = 0: Grid2.FixedCols = 0             'Numero de Columnas fijas
      Grid1.FixedRows = 1: Grid2.FixedRows = 1              'Numero de Filas Fijas (Titulos)
      Grid1.Rows = 1: Grid2.Rows = 1                  '1 Fila
      Grid1.Cols = UBound(Campos) + 1
      Grid2.Cols = 5
    End If
    ' Control de Fila a utilizar, se añade si es necesario
    If Grid1.Rows <= fila Then Grid1.Rows = fila + 1: If Grid2.Rows <= fila Then Grid2.Rows = fila + 1
    ' Situamos una a una las Columnas.
    For Columna = 0 To UBound(Campos)
        Grid1.TextMatrix(fila, Columna) = Campos(Columna)
    Next
    DoEvents
    If fila > 0 Then
        CodigoArt = Campos(16)      'campo 16 = SKU
        CodigoArt = Trim(CodigoArt)
        If Len(CodigoArt) = 5 Then
            CodigoArt = "0" & CodigoArt
        End If
        s = "SELECT * FROM VALORSTK WHERE codart LIKE ""%" & CodigoArt & "%"";"
        tRs.Open s, cnn, adOpenDynamic, adLockOptimistic 'adLockOptimistic
        With tRs
            If (.EOF = True) And (.BOF = True) Then
                ' Si no hay datos...
                MsgBox "NADA ups!"
            Else
                n = 0
                Do While Not .EOF
                    n = n + 1
                    'MsgBox tRs.Fields("codart") & " - $" & tRs.Fields("nombre") & " - $" & tRs.Fields("precio") & " - " & tRs.Fields("exis")
                    Grid2.TextMatrix(fila, 0) = tRs.Fields("codart")
                    Grid2.TextMatrix(fila, 1) = tRs.Fields("nombre")
                    mdbnum = CDbl(tRs.Fields("precio"))
                    Mdbprecio = Format(mdbnum, "0.00")
                    Grid2.TextMatrix(fila, 2) = Mdbprecio
                    Grid2.TextMatrix(fila, 3) = tRs.Fields("exis")
                    .MoveNext
                Loop
            End If
        End With
        tRs.Close
    Else
        Grid2.TextMatrix(0, 0) = "codart"
        Grid2.TextMatrix(0, 1) = "nombre"
        Grid2.TextMatrix(0, 2) = "precio"
        Grid2.TextMatrix(0, 3) = "exis"
    End If

    ' Aumentamos número de Fila
    fila = fila + 1
    'prgb1.Value = fila
    
  Wend
  
  Close #Fichero
  Grid1.Redraw = True: Grid2.Redraw = True
  'prgb1.Visible = False
  Command1.Enabled = True
  Command2.Enabled = True
  Command3.Enabled = True
  Label4.Visible = False
  
    'tRs.Close
    cnn.Close
    
End Sub
   
Private Sub Check1_Click()

If Check1.Value = 1 Then
    ShowLOG = True
Else
    ShowLOG = False
End If

End Sub

Private Sub ChkBarras_Click()

'MsgBox ChkBarras.Value

End Sub

'BOTON IMPORTAR
'CARGAMOS LOS DATOS DE LA PRIMARY DATABASE Y CARGAMOS EL CSV PARA UNA COMPARACION MANUAL O VERIFICACION
'------------------------------------------------------------------------------------------------------
Private Sub Command1_Click()

Dim ConverTx As String

Label1.Caption = "Productos IBERICO - Archivo: ...\data\database2.mdb"

On Error Resume Next
EstCMD.InitDir = App.Path & "\data\"
EstCMD.Filter = "Archivo CSV (*.csv)|*.csv|Archivos CSV"
EstCMD.DialogTitle = "Seleccione el archivo CSV DE TIENDA NUBE..."
EstCMD.CancelError = True
EstCMD.ShowOpen

If err.Number = 32755 Then Exit Sub

ConverTx = EstCMD.FileName

Label2.Caption = "Productos TIENDA NUBE - Archivo:" & ConverTx

Command2.Enabled = False
DoEvents
CargarCSV Grilla1, Grilla2, App.Path & "\data\database2.mdb", ConverTx
Command2.Enabled = True

End Sub

'BOTON EXPORTAR
'CARGAMOS LOS DATOS DE LA PRIMARY DATABASE Y CARGAMOS EL CSV PRIMARIO Y LUEGO UNIFICAMOS LOS DATOS
'Y EXPORTAMOS EN UN NUEVO CSV.
'------------------------------------------------------------------------------------------------------
Private Sub Command2_Click()

Dim ConverTx As String, ConvertXX As String

'abrir CSV origen
'On Error Resume Next
EstCMD.InitDir = App.Path & "\data\"
EstCMD.Filter = "Archivo CSV (*.csv)|*.csv|Archivos CSV"
EstCMD.DialogTitle = "Seleccione el archivo CSV DE TIENDA NUBE..."
EstCMD.CancelError = True
EstCMD.ShowOpen

If err.Number = 32755 Then Exit Sub

ConverTx = EstCMD.FileName

If Len(ConverTx) > 15 Then
    Label2.Caption = "Productos TIENDA NUBE - Archivo: " & "..." & Right$(ConverTx, 15)
Else
    Label2.Caption = "Productos TIENDA NUBE - Archivo: " & ConverTx
End If

'preguntar CSV destino
'On Error Resume Next
EstCMDS.InitDir = App.Path & "\data\"
EstCMDS.Filter = "Archivo CSV (*.csv)|*.csv|Archivos CSV"
EstCMDS.DialogTitle = "Guardar archivo CSV ACTUALIZADO de TIENDA NUBE como..."
EstCMDS.CancelError = True
EstCMDS.ShowSave

If err.Number = 32755 Then Exit Sub

ConvertXX = EstCMDS.FileName

'Label2.Caption = "Productos TIENDA NUBE - Archivo:" & ConvertTx

Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False

DoEvents
FrmSearching.Show , Me

If ChkInternal.Value = 1 Then   'utilizar base de datos interna
    Label1.Caption = "Productos IBERICO - Archivo: ...\Database\L_L_Stock.llk"
    ExportarCSVfromDB Grilla1, Grilla2, ConverTx, ConvertXX
Else
    If ChkExternal.Value = 1 Then   'utilizar base de datos externa o MDB
        ExportarCSVfromMDB Grilla1, Grilla2, App.Path & "\data\database2.mdb", ConverTx, ConvertXX
        Label1.Caption = "Productos IBERICO - Archivo: ...\data\database2.mdb"
    Else
        MsgBox "Por favor seleccione que tipo de BASE desea utilizar para Actualizar."
        Exit Sub
    End If
End If

Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True

Unload FrmSearching

MsgBox "Archivo CSV Exportado Correctamente como: " & ConvertXX

End Sub

Private Sub Command3_Click()

Unload Me

End Sub

Private Sub Form_Load()

'/// mostrar LOG??
If Check1.Value = 1 Then
    ShowLOG = True
Else
    ShowLOG = False
End If

End Sub

Private Sub Grilla1_Scroll()

Grilla2.Row = Grilla1.Row
Grilla2.Col = Grilla1.Col
Grilla2.TopRow = Grilla1.TopRow
Grilla2.LeftCol = Grilla1.LeftCol

End Sub

Private Sub Grilla1_SelChange()

'Grilla2.Row = Grilla1.Row
'Grilla2.Col = Grilla1.Col
'Grilla2.RowSel = Grilla1.Row
'Grilla2.ColSel = Grilla1.Col

End Sub

Private Sub Grilla2_Scroll()

Grilla1.Row = Grilla2.Row
Grilla1.Col = Grilla2.Col
Grilla1.TopRow = Grilla2.TopRow
Grilla1.LeftCol = Grilla2.LeftCol

End Sub

Private Sub Grilla2_SelChange()

Grilla1.Row = Grilla2.Row
Grilla1.Col = Grilla2.Col
Grilla1.RowSel = Grilla2.RowSel
Grilla1.ColSel = Grilla2.ColSel

End Sub
