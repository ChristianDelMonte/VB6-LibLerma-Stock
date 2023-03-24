VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form IngVCons 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Consulta / Informes de datos sobre Bajas ingresadas en sistema."
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8760
   Icon            =   "IngVCons.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView Lv1 
      Height          =   5625
      Left            =   90
      TabIndex        =   13
      Top             =   1680
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   9922
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   945
      MaxLength       =   20
      TabIndex        =   12
      Text            =   "0"
      Top             =   7470
      Width           =   1140
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Rendir"
      Height          =   375
      Left            =   2205
      TabIndex        =   10
      Top             =   7425
      Width           =   1230
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   6300
      TabIndex        =   9
      Top             =   7425
      Width           =   1230
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   7605
      TabIndex        =   1
      ToolTipText     =   "Cerrar y volver"
      Top             =   7425
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opciones "
      Height          =   735
      Left            =   90
      TabIndex        =   0
      Top             =   855
      Width           =   8565
      Begin VB.CommandButton Command1 
         Caption         =   "Consultar"
         Height          =   375
         Left            =   7335
         TabIndex        =   8
         Top             =   225
         Width           =   1140
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Rendidos"
         Height          =   195
         Left            =   2115
         TabIndex        =   7
         Top             =   450
         Width           =   1230
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Sin Rendir"
         Height          =   240
         Left            =   2115
         TabIndex        =   6
         Top             =   180
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   990
         TabIndex        =   2
         Text            =   "1"
         Top             =   270
         Width           =   780
      End
      Begin VB.Label Label2 
         Caption         =   "Sucursal:"
         Height          =   255
         Left            =   135
         TabIndex        =   3
         Top             =   315
         Width           =   750
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   8730
      TabIndex        =   4
      Top             =   0
      Width           =   8760
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "BAJAS | CONSULTAS"
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
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   8700
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Comp. N°:"
      Height          =   195
      Left            =   135
      TabIndex        =   11
      Top             =   7515
      Width           =   735
   End
End
Attribute VB_Name = "IngVCons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub DeleteFile(FilePath As String)
On Error GoTo error
Kill FilePath$
Exit Sub
error:  MsgBox err.Description, vbExclamation, "Error"
End Sub

Function Create_Recordset(ListView As ListView) As ADODB.Recordset

On Error GoTo error_function
    
    Dim iRow As Long
    Dim item As ListItem
    Dim iCol As Long
    
    ' Instanciar y crear el recordset sin conexión
    Dim rs As Recordset
    Set rs = New Recordset
    
    ' Añadir los campos con el método Append del objeto Field de ADO
    With ListView
        For iCol = 1 To ListView.ColumnHeaders.Count
            rs.Fields.Append .ColumnHeaders(iCol).Text, adVarChar, 250
        Next
    End With
    
    ' Abrir el Recordset para añadir los items con le método AddNew
    rs.Open
    With rs
        ' recorrer el control
        For iRow = 1 To ListView.ListItems.Count
            Set item = ListView.ListItems(iRow)
            
            .AddNew ' Crear nuevo
            
            ' asignar el valor ( El item )
            .Fields(0).Value = item.Text
            
            'asignar Demás valores ( Los SubItems )
            For iCol = 1 To ListView.ColumnHeaders.Count - 1
                .Fields(iCol).Value = item.SubItems(iCol)
            Next
        Next
    
    End With
    ' retornar el rec a la función
    Set Create_Recordset = rs
    
Exit Function

' Errores
error_function:

MsgBox err.Description, vbCritical

End Function

Private Sub Command1_Click()

Dim NumSucursal As Long, UltimoReg As Long, Contador As Long
Dim ConsultaVentas As VentasData, SubElemento As ListItem

'//reiniciamos el listview
Lv1.ListItems.Clear

NumSucursal = CLng(Trim(Combo1.Text))

If Option1.Value = True Then
    UltimoReg = GetVentasLastReg(NumSucursal)
    For Contador = 1 To UltimoReg - 1
        ConsultaVentas = BuscarVentas(Contador, NumSucursal, 0)
        'Lv1.ListItems.Add , , Trim(ConsultaVentas.VIDesc)
        Set SubElemento = Lv1.ListItems.Add(, , (Trim(ConsultaVentas.Vdate)))
        SubElemento.SubItems(1) = Trim(ConsultaVentas.VSuc)
        SubElemento.SubItems(2) = Trim(ConsultaVentas.VTurno)
        SubElemento.SubItems(3) = Trim(ConsultaVentas.VICode)
        SubElemento.SubItems(4) = Trim(ConsultaVentas.VIDesc)
        SubElemento.SubItems(5) = Trim(ConsultaVentas.VICant)
        If CLng(ConsultaVentas.VRendido) = 0 Then
            SubElemento.SubItems(6) = "NO"
        Else
            SubElemento.SubItems(6) = "SI"
        End If
    Next
ElseIf Option2.Value = True Then
    UltimoReg = GetRendLastReg(NumSucursal)
    For Contador = 1 To UltimoReg - 1
        ConsultaVentas = BuscarVentas(Contador, NumSucursal, 1)
        'Lv1.ListItems.Add , , Trim(ConsultaVentas.VIDesc)
        Set SubElemento = Lv1.ListItems.Add(, , (Trim(ConsultaVentas.Vdate)))
        SubElemento.SubItems(1) = Trim(ConsultaVentas.VSuc)
        SubElemento.SubItems(2) = Trim(ConsultaVentas.VTurno)
        SubElemento.SubItems(3) = Trim(ConsultaVentas.VICode)
        SubElemento.SubItems(4) = Trim(ConsultaVentas.VIDesc)
        SubElemento.SubItems(5) = Trim(ConsultaVentas.VICant)
        If CLng(ConsultaVentas.VRendido) = 0 Then
            SubElemento.SubItems(6) = "NO"
        Else
            SubElemento.SubItems(6) = "SI"
        End If
        SubElemento.SubItems(7) = Trim(ConsultaVentas.VROp)
    Next
End If

End Sub

Private Sub Command2_Click()

    Dim rs As ADODB.Recordset
    Dim Nombre_seccion As String
    
    ' llamar la función Create_Recordset
    Set rs = Create_Recordset(Lv1)
    
    If Not rs Is Nothing Then
        
       'Indicar en esta variable el nombre de la sección en la que se encuentran los rptTextBox para cada campo
        Nombre_seccion = "Sección1"
        
        'Asignarle a los textbox del datareport, los DataField que corresponden a los nombres de encabezados
        With DataReport1
                 
            .Sections(Nombre_seccion).Controls.item("Texto1").DataField = "FECHA"
            .Sections(Nombre_seccion).Controls.item("Texto2").DataField = "SUC."
            .Sections(Nombre_seccion).Controls.item("Texto3").DataField = "TURNO"
            .Sections(Nombre_seccion).Controls.item("Texto4").DataField = "SKU"
            .Sections(Nombre_seccion).Controls.item("Texto5").DataField = "ARTICULO"
            .Sections(Nombre_seccion).Controls.item("Texto6").DataField = "CANT."
            .Sections(Nombre_seccion).Controls.item("Texto7").DataField = "REND?"
            
            ' cambiar el caption del label Titulo
            '.Sections("Sección4").Controls.item("lbltitulo").Caption = "Informe de formulario: " & Me.Caption
            
            ' Asignarle al datasource el origen de datos, es decir el recordset que devolvió la función Create_Recordset
            Set .DataSource = rs
            
            'Cargar y muestrar el informe
            .Show vbModal
            
            ' Liberar los recursos
            If rs.State = adStateOpen Then rs.Close
            Set rs = Nothing
            
        End With
    End If
    
End Sub

Private Sub Command3_Click()

Unload Me

End Sub

Private Sub Command4_Click()

Dim NumSucursal As Long, UltimoReg As Long, UltimoRReg As Long, Contador As Long
Dim ConsultaVentas As VentasData, SubElemento As ListItem
Dim RndVentas As VentasData, Resultado As Boolean, Fln As String

If Trim(Text1.Text) = "" Or Trim(Text1.Text) = "0" Then
    MsgBox "Para Rendir debera ingresar el numero de comprobante correspondiente."
    Exit Sub
End If

'//reiniciamos el listview
Lv1.ListItems.Clear

NumSucursal = CLng(Trim(Combo1.Text))

'ultimo registro de ventas
UltimoReg = GetVentasLastReg(NumSucursal)
'ultimo registro de rendidos
'UltimoRReg = GetRendLastReg(NumSucursal)

For Contador = 1 To UltimoReg - 1
    ConsultaVentas = BuscarVentas(Contador, NumSucursal)
    'Lv1.ListItems.Add , , Trim(ConsultaVentas.VIDesc)
    Set SubElemento = Lv1.ListItems.Add(, , (Trim(ConsultaVentas.Vdate)))
        RndVentas.Vdate = Trim(ConsultaVentas.Vdate)
    SubElemento.SubItems(1) = Trim(ConsultaVentas.VSuc)
        RndVentas.VSuc = Trim(ConsultaVentas.VSuc)
    SubElemento.SubItems(2) = Trim(ConsultaVentas.VTurno)
        RndVentas.VTurno = Trim(ConsultaVentas.VTurno)
    SubElemento.SubItems(3) = Trim(ConsultaVentas.VICode)
        RndVentas.VICode = Trim(ConsultaVentas.VICode)
    SubElemento.SubItems(4) = Trim(ConsultaVentas.VIDesc)
        RndVentas.VIDesc = Trim(ConsultaVentas.VIDesc)
    SubElemento.SubItems(5) = Trim(ConsultaVentas.VICant)
        RndVentas.VICant = Trim(ConsultaVentas.VICant)
    SubElemento.SubItems(6) = "SI"
        RndVentas.VRendido = 1
        RndVentas.VICosto = Trim(ConsultaVentas.VICosto)
        RndVentas.VRdate = Date$
        RndVentas.VROp = Trim(Text1.Text)
    SubElemento.SubItems(7) = Text1.Text
    If RendirVentas(RndVentas, 0) = False Then
        MsgBox "Oups!"
    End If
Next

'una vez rendido se elimina el archivo de origen
Fln = "\Data\vlts\data" & "-" & NumSucursal & VentasExt
Call DeleteFile(App.Path & Fln)

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

Private Sub Option1_Click()

'bajas sin rendir
If Option1.Value = True Then
    Text1.Enabled = True
    Command4.Enabled = True
Else
    Text1.Enabled = False
    Command4.Enabled = False
End If

End Sub

Private Sub Option2_Click()

'bajas rendidas
If Option2.Value = True Then
    Text1.Enabled = False
    Command4.Enabled = False
Else
    Text1.Enabled = True
    Command4.Enabled = True
End If

End Sub
