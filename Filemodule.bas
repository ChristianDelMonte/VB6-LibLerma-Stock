Attribute VB_Name = "FileModule"

'***********************************************************************************************
'** Prometheus - Libreria Lerma Stock
'** es una aplicacion desarrollada exclusivamente para llevar a cabo las tareas relacionadas
'** con el control de stock de la libreria. Control de Vueltos y otras tareas necesarias
'** que dependen de mi exclusiva responsabilidad.
'** -------------------------------------------------------------------------------------------
'** Desarrollado por: Christian Adrian Del Monte
'** Dudas o consultas: creadig@gmail.com / creadig@hotmail.com
'** Ultima actualizacion: Agosto 2019
'***********************************************************************************************
Public Type STKDatabase
    Aid As Long
    CodeSec As String * 20
    ArtName As String * 50
    Exis As String * 6
    Precio As String * 10
End Type

Public Type VentasData              'tipo para controlar los articulos de vuelto en ventas
    Vid As Integer                  'identificador de registro o numero de id (unico)
    Vdate As String * 10            'fecha de carga de datos o fecha de venta o baja del aticulo
    VSuc As Long                    'numero de sucursal involucrada en la transaccion
    VTurno As Integer               'turno responsable (1=mañana - 2=tarde)
    VICode As String * 8            'codigo del articulo
    VIDesc As String * 100          'descripcion del articulo
    VICosto As String * 10          'precio del articulo
    VICant As Integer               'Cantidad de articulos
    VRendido As Integer             'articulo rendido? dado de baja? 1=si 0=no
    VRdate As String * 10           'fecha de rendicion
    VROp As String * 20             'numero de comprobante de operacion
End Type

Public cnn As ADODB.Connection
Public ControlVentas As VentasData  'para el control de vueltos de ventas
Public StockDatabase As STKDatabase 'base de datos general de la aplicacion
Public Const InvFilename = "\Database\L_L_Stock.llk"
Public Const VentasExt = ".inv"
Public Const RendVentasExt = ".ipk"

'// ***************************************************************************************************
'/// Funcion para consultar de la base de inventario de articulos
'/// REVISAR REVISAR REVISAR REVISAR REVISAR REVISAR REVISAR
Public Function CargaInventario(Wid As Long) As STKDatabase

Dim LastReg As Long, i As Long

'/// abrimos el archivo de productos
'On Error GoTo err
Open App.Path & InvFilename For Random As #16 Len = Len(StockDatabase)

'/// check for the ID to load
If Wid <> 0 Then
    LastReg = Wid
Else
    GoTo err
End If

'/// cargamos
Get #16, LastReg, StockDatabase

CargaInventario.Aid = StockDatabase.Aid
CargaInventario.ArtName = Trim(StockDatabase.ArtName)
CargaInventario.CodeSec = Trim(StockDatabase.CodeSec)
CargaInventario.Exis = Trim(StockDatabase.Exis)
CargaInventario.Precio = Trim(StockDatabase.Precio)

Close #16

Exit Function

'/// if there is an error ------------------------------------------
err:
'WriteErrors "Error en SaveINVData > Module INV. - " & err.Number & " - " & err.Description
Close #16
StockDatabase.Aid = -1
CargaInventario.Aid = StockDatabase.Aid

End Function

'// ***************************************************************************************************
'/// Funcion para actualizar y crear la base de inventario de articulos
'/// todos los articulos de la libreria en stock deben ir aqui.
Public Function GuardaInventario(Dato As STKDatabase, WOptionalID As Long) As Boolean

Dim LastReg As Long

'/// abrimos el archivo de productos
On Error GoTo err
Open App.Path & InvFilename For Random As #12 Len = Len(Dato)

'/// check for the ID to save
If WOptionalID = 0 Or WOptionalID = -1 Then
    LastReg = LOF(12) \ Len(STKDatabase)
    LastReg = LastReg + 1
Else
    LastReg = WOptionalID
End If

'/// seteamos los datos del inventario a guardar
StockDatabase.Aid = LastReg
StockDatabase.ArtName = Trim(Dato.ArtName)
StockDatabase.CodeSec = Trim(Dato.CodeSec)
StockDatabase.Exis = Trim(Dato.Exis)
StockDatabase.Precio = Trim(Dato.Precio)
'/// guardamos
Put #12, LastReg, StockDatabase
Close #12

GuardaInventario = True
Exit Function

'/// if there is an error ------------------------------------------
err:
'WriteErrors "Error en SaveINVData > Module INV. - " & err.Number & " - " & err.Description
Close #12
GuardaInventario = False

End Function

'// ***************************************************************************************************
Public Function GuardarVentas(Articulo As VentasData, WOptionalID As Integer) As Boolean

Dim FileName As String
Dim LastReg As Integer

FileName = "\Data\vlts\data" & "-" & Articulo.VSuc & VentasExt

'/// abrimos el archivo de productos
On Error GoTo err
Open App.Path & FileName For Random As #12 Len = Len(ControlVentas)

'/// check for the ID to save
If WOptionalID = 0 Or WOptionalID = -1 Then
    LastReg = LOF(12) \ Len(ControlVentas)
    LastReg = LastReg + 1
Else
    LastReg = WOptionalID
End If

'/// seteamos los datos del inventario a guardar
ControlVentas.Vid = LastReg
ControlVentas.Vdate = Articulo.Vdate
ControlVentas.VSuc = Articulo.VSuc
ControlVentas.VTurno = Articulo.VTurno
ControlVentas.VICode = Trim(Articulo.VICode)
ControlVentas.VIDesc = Trim(Articulo.VIDesc)
ControlVentas.VICosto = Trim(Articulo.VICosto)
ControlVentas.VICant = Articulo.VICant
ControlVentas.VRendido = 0
ControlVentas.VRdate = "000000"
ControlVentas.VROp = "000000"
'/// guardamos
Put #12, LastReg, ControlVentas
Close #12

GuardarVentas = True
Exit Function

'/// if there is an error ------------------------------------------
err:
'WriteErrors "Error en SaveINVData > Module INV. - " & err.Number & " - " & err.Description
Close #12
GuardarVentas = False

End Function

'// ***************************************************************************************************
'/// funcion para buscar un articulo dentro de la base de ventas
Public Function BuscarVentas(RegId As Long, SucId As Long, Optional Rendido As Long) As VentasData

Dim LastReg As Long
Dim FileName As String
Dim BC As String

If Rendido = 1 Then
    FileName = "\Data\vlts\rends\data" & "-" & SucId & RendVentasExt
    LastReg = GetRendLastReg(SucId)
Else
    FileName = "\Data\vlts\data" & "-" & SucId & VentasExt
    LastReg = GetVentasLastReg(SucId)
End If

'/// abrimos el archivo de los abonados
On Error GoTo err
Open App.Path & FileName For Random As #14 Len = Len(ControlVentas)

If RegId < LastReg Then
    Get #14, RegId, ControlVentas
    BuscarVentas.Vdate = Trim(ControlVentas.Vdate)
    BuscarVentas.VICant = ControlVentas.VICant
    BuscarVentas.VICode = Trim(ControlVentas.VICode)
    BuscarVentas.VICosto = Trim(ControlVentas.VICosto)
    BuscarVentas.Vid = ControlVentas.Vid
    BuscarVentas.VIDesc = Trim(ControlVentas.VIDesc)
    BuscarVentas.VSuc = ControlVentas.VSuc
    BuscarVentas.VTurno = ControlVentas.VTurno
    BuscarVentas.VRdate = Trim(ControlVentas.VRdate)
    BuscarVentas.VRendido = ControlVentas.VRendido
    BuscarVentas.VROp = Trim(ControlVentas.VROp)
Else
    BuscarVentas.Vid = -1
    BuscarVentas.VIDesc = "Error"
End If

Close #14
Exit Function

'/// if there is an error ------------------------------------------
err:
'WriteErrors "Error en GetAGEData > Module ABN. - " & err.Number & " - " & err.Description
BuscarVentas.Vid = -1
BuscarVentas.VIDesc = "Error"
Close #14

End Function

'// ***************************************************************************************************
Public Function RendirVentas(Articulo As VentasData, WOptionalID As Integer) As Boolean

Dim FileName As String
Dim LastReg As Integer

FileName = "\Data\vlts\rends\data" & "-" & Articulo.VSuc & RendVentasExt

'/// abrimos el archivo de productos
On Error GoTo err
Open App.Path & FileName For Random As #17 Len = Len(ControlVentas)

'/// check for the ID to save
If WOptionalID = 0 Or WOptionalID = -1 Then
    LastReg = LOF(17) \ Len(ControlVentas)
    LastReg = LastReg + 1
Else
    LastReg = WOptionalID
End If

'/// seteamos los datos del inventario a guardar
ControlVentas.Vid = LastReg
ControlVentas.Vdate = Trim(Articulo.Vdate)
ControlVentas.VSuc = Articulo.VSuc
ControlVentas.VTurno = Articulo.VTurno
ControlVentas.VICode = Trim(Articulo.VICode)
ControlVentas.VIDesc = Trim(Articulo.VIDesc)
ControlVentas.VICosto = Trim(Articulo.VICosto)
ControlVentas.VICant = Articulo.VICant
ControlVentas.VRendido = Articulo.VRendido
ControlVentas.VRdate = Trim(Articulo.VRdate)
ControlVentas.VROp = Trim(Articulo.VROp)
'/// guardamos
Put #17, LastReg, ControlVentas
Close #17

RendirVentas = True
Exit Function

'/// if there is an error ------------------------------------------
err:
'WriteErrors "Error en SaveINVData > Module INV. - " & err.Number & " - " & err.Description
Close #17
RendirVentas = False

End Function

'// ***************************************************************************************************
'/// funcion para buscar un articulo dentro de la base de datos
Public Function BuscarArticulo2(CodigoArt As Long) As STKDatabase

Dim LastReg As Long
Dim BC As String

'/// abrimos el archivo de los abonados
On Error GoTo err
Open App.Path & InvFilename For Random As #14 Len = Len(StockDatabase)

LastReg = GetInvLastReg
If CodigoArt < LastReg Then
    Get #14, CodigoArt, StockDatabase
    BuscarArticulo2.ArtName = Trim(StockDatabase.ArtName)
    BuscarArticulo2.Precio = Trim(StockDatabase.Precio)
    BuscarArticulo2.Aid = StockDatabase.Aid
    BuscarArticulo2.CodeSec = Trim(StockDatabase.CodeSec)
    BuscarArticulo2.Exis = Trim(StockDatabase.Exis)
Else
    BuscarArticulo2.Aid = -1
    BuscarArticulo2.ArtName = "Error"
End If

Close #14
Exit Function

'/// if there is an error ------------------------------------------
err:
'WriteErrors "Error en GetAGEData > Module ABN. - " & err.Number & " - " & err.Description
BuscarArticulo2.ArtName = "Error"
Close #14

End Function

'// ***************************************************************************************************
'// funcion para extraer de un archivo los datos de creacion y ultima modificacion
Public Function GetFechaArch(Path As String) As String
  
    'Variable de tipo FileSystemObject y File
    Dim o_Fso As New FileSystemObject
    Dim Archivo As File
  
    ' Lee las propiedades del archivo mediante GetFile
    On Error GoTo err
    Set Archivo = o_Fso.GetFile(Path)

    'Visualiza el resultado: Creación ,acceso y modificado etc..
    'MsgBox "Fecha de creación del archivo: " & Format(Archivo.DateCreated), vbInformation
    'MsgBox "Fecha de modificación : " & Format(Archivo.DateLastModified), vbInformation
    GetFechaArch = Format(Archivo.DateLastModified)
    'MsgBox "Fecha de del último acceso: " & Format(Archivo.DateLastAccessed), vbInformation
    'MsgBox "Tamaño del archivo : " & Format(Archivo.Size) & " Bytes", vbInformation
    'MsgBox "Tipo de archivo : " & Format(Archivo.Type), vbInformation
      
    ' Elimina las variables de objeto
    Set Archivo = Nothing
    Set o_Fso = Nothing
    Exit Function

'/// if there is an error ------------------------------------------
err:
    Set Archivo = Nothing
    Set o_Fso = Nothing
GetFechaArch = "Inexistente"

End Function

'// ***************************************************************************************************
'/// Funcion para extraer el ultimo registro de la base de inventario de articulos
Public Function GetInvLastReg() As Long

Dim LastReg As Long

'/// abrimos el archivo de productos
'On Error GoTo err
Open App.Path & InvFilename For Random As #12 Len = Len(StockDatabase)

'/// check for the last ID
LastReg = LOF(12) \ Len(StockDatabase)
LastReg = LastReg + 1

Close #12

GetInvLastReg = LastReg
Exit Function

'/// if there is an error ------------------------------------------
err:
'WriteErrors "Error en SaveINVData > Module INV. - " & err.Number & " - " & err.Description
Close #12
GetInvLastReg = -1

End Function

'// ***************************************************************************************************
'/// Funcion para extraer el ultimo registro de la base de ventas de articulos
Public Function GetVentasLastReg(NumSuc As Long) As Long

Dim LastReg As Long, FileName As String

FileName = "\Data\vlts\data" & "-" & NumSuc & VentasExt

'/// abrimos el archivo de productos
'On Error GoTo err
Open App.Path & FileName For Random As #11 Len = Len(ControlVentas)

'/// check for the last ID
LastReg = LOF(11) \ Len(ControlVentas)
LastReg = LastReg + 1

Close #11

GetVentasLastReg = LastReg
Exit Function

'/// if there is an error ------------------------------------------
err:
'WriteErrors "Error en SaveINVData > Module INV. - " & err.Number & " - " & err.Description
Close #11
GetVentasLastReg = -1

End Function

'// ***************************************************************************************************
'/// Funcion para extraer el ultimo registro de la base de ventas de articulos
Public Function GetRendLastReg(NumSuc As Long) As Long

Dim LastReg As Long, FileName As String

FileName = "\Data\vlts\rends\data" & "-" & NumSuc & RendVentasExt

'/// abrimos el archivo de productos
'On Error GoTo err
Open App.Path & FileName For Random As #11 Len = Len(ControlVentas)

'/// check for the last ID
LastReg = LOF(11) \ Len(ControlVentas)
LastReg = LastReg + 1

Close #11

GetRendLastReg = LastReg
Exit Function

'/// if there is an error ------------------------------------------
err:
'WriteErrors "Error en SaveINVData > Module INV. - " & err.Number & " - " & err.Description
Close #11
GetRendLastReg = -1

End Function

'// ***************************************************************************************************
'// funcion para buscar la descripcion de un articulo por su codigo principal o secundario (de barras)
'// retorna la descripcion del articulo buscado o "ERR" en caso de no encontrar nada
'// ***************************************************************************************************
Public Function BuscarArticuloMDB(PrimaryMDB As String, CodigoArt As String) As String

    Dim s As String
    Dim tRs As ADODB.Recordset
    Dim n As Long
    
    Set cnn = New ADODB.Connection
    Provider = "Microsoft.Jet.OLEDB.4.0"
    cnn.Open "Provider=" & Provider & "; Data Source=" & PrimaryMDB

    Set tRs = New ADODB.Recordset
    
    DoEvents
     
    CodigoArt = Trim(CodigoArt)
    If Len(CodigoArt) = 5 Then
        CodigoArt = "0" & CodigoArt
    End If
    
    s = "SELECT * FROM VALORSTK WHERE codart LIKE ""%" & CodigoArt & "%"";"
    tRs.Open s, cnn, adOpenDynamic, adLockOptimistic 'adLockOptimistic
    
    With tRs
        If (.EOF = True) And (.BOF = True) Then
            ' Si no hay datos...
            'MsgBox "No se encontro el Articulo"
            BuscarArticuloMDB = "NO ENCONTRADO"
            tRs.Close
            cnn.Close
            Exit Function
        Else
            n = 0
            Do While Not .EOF
                n = n + 1
                 BuscarArticuloMDB = Trim(tRs.Fields("nombre"))
                .MoveNext
            Loop
        End If
    End With
    
    tRs.Close
'tRs.Close
cnn.Close

End Function

