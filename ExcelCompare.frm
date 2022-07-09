VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Comparacion de planillas"
   ClientHeight    =   6660
   ClientLeft      =   2985
   ClientTop       =   2265
   ClientWidth     =   13665
   Icon            =   "ExcelCompare.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6660
   ScaleWidth      =   13665
   Begin VB.CheckBox chkCase 
      BackColor       =   &H8000000C&
      Caption         =   "Ignorar caso"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3210
      TabIndex        =   17
      Top             =   720
      Width           =   1575
   End
   Begin VB.CheckBox chkShow 
      BackColor       =   &H8000000C&
      Caption         =   "Mostrar resultado"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1290
      TabIndex        =   16
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   6840
      TabIndex        =   15
      Top             =   960
      Width           =   6735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   1260
      TabIndex        =   14
      Top             =   960
      Width           =   5445
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Cargar"
      Height          =   495
      Index           =   1
      Left            =   10920
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox cmbSheet 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   6840
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   360
      Width           =   6735
   End
   Begin VB.TextBox txtFile 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   6840
      TabIndex        =   9
      Top             =   50
      Width           =   6135
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   13050
      TabIndex        =   8
      Top             =   45
      Width           =   495
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H8000000C&
      Caption         =   "Formulas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   1000
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H8000000C&
      Caption         =   "Valores"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Cargar"
      Height          =   495
      Index           =   0
      Left            =   4800
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox cmbSheet 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   360
      Width           =   5775
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   6300
      TabIndex        =   1
      Top             =   45
      Width           =   435
   End
   Begin VB.TextBox txtFile 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   50
      Width           =   6105
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   5175
      Index           =   1
      Left            =   6840
      TabIndex        =   7
      Top             =   1350
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   9128
      _Version        =   393216
      FixedRows       =   0
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   5175
      Index           =   0
      Left            =   0
      TabIndex        =   12
      Top             =   1320
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   9128
      _Version        =   393216
      FixedRows       =   0
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblResults 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      Caption         =   "Listo..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4860
      TabIndex        =   13
      Top             =   735
      Width           =   5055
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000C&
      Caption         =   "Hoja N"
      Height          =   195
      Left            =   300
      TabIndex        =   2
      Top             =   420
      Width           =   585
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Program to Compare 2 excel files and highlight the differences
'submitted by Tom Pydeski
'
'The Original example is at :
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=61399&lngWId=1
'and was submitted by Cristiano Couto.
'I modified it to open 2 excel files and compare each cell.
'
'If the exe file does not run, copy the MSFLXGRD.xxx files to the
'c:\windows\system32 directory.
'The ocx must be registered by the following:
'""Regsvr32.exe MSFLXGRD.ocx"
'
'Eventually the code in the .frm file can be converted to run within excel as a vb macro.
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Dim Col1st As Integer
Dim Confirm As Long
Dim eTitle$
Dim EMess$
Dim mError As Long
Dim Inits As Byte
Dim r As Integer
Dim c As Integer
Dim i As Integer
Dim Changing As Byte
Dim InitDir$
Dim lRet As Long

Private Sub chkCase_Click()
'if we changed our case selection, redo the compare
Compare
End Sub

Private Sub chkShow_Click()
'if our check box to show results is selected, then load the results form
If chkShow.Value = vbChecked Then
    frmResults.Show
Else
    Unload frmResults
End If
End Sub

Private Sub cmbSheet_Click(Index As Integer)
'below keeps this event from firing again during this procedure
If Changing = 1 Then Exit Sub
'set the variable to indicate the change is in progress
Changing = 1
'load the sheet
cmdLoad_Click (Index)
'if we have already loaded the form, then change the opposite combo to match
If Inits = 1 Then
    cmbSheet(1 - Index).ListIndex = cmbSheet(Index).ListIndex
    DoEvents
    cmdLoad_Click (1 - Index)
    Compare
End If
DoEvents
Refresh
Changing = 0
End Sub

Private Sub cmdOpen_Click(Index As Integer)
Dim OFName As OPENFILENAME
Dim XLApp As Object
Dim Wrk As Object
Dim Sht As Object
OFName.lStructSize = Len(OFName)
'Set the parent window
OFName.hwndOwner = Me.hWnd
'Set the application's instance
OFName.hInstance = App.hInstance
'Select a filter
OFName.lpstrFilter = "Excel Files (*.xls)" + Chr$(0) + "*.xls" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
'create a buffer for the file
OFName.lpstrFile = Space$(254)
'set the maximum length of a returned file
OFName.nMaxFile = 255
'Create a buffer for the file title
OFName.lpstrFileTitle = Space$(254)
'Set the maximum length of a returned file title
OFName.nMaxFileTitle = 255
'Set the initial directory if we have one saved
If Len(InitDir$) > 0 Then
    OFName.lpstrInitialDir = InitDir$
End If
'Set the title
OFName.lpstrTitle = "Open XLS File"
'No flags
OFName.flags = 0
'Show the 'Open File'-dialog
If GetOpenFileName(OFName) Then
    'get the filename and full path
    txtFile(Index).Text = Trim$(OFName.lpstrFile)
    'get the filename only (no path)
    Text1(Index).Text = Trim$(OFName.lpstrFileTitle)
    'extract the path from the full filename
    InitDir$ = Left$(txtFile(Index).Text, Len(txtFile(Index).Text) - Len(Text1(Index).Text))
    'save our path for the next time
    SaveSetting App.EXEName, "Settings", "InitDir", InitDir$
    'clear the combo box for the sheet list
    cmbSheet(Index).Clear
    'Create a new instance of Excel
    Set XLApp = CreateObject("Excel.Application")
    'Open the XLS file. The two parameters representes, UpdateLink = False and ReadOnly = True.
    'These parameters have this setting so they don't occur any error on broken links and allready opened XLS file.
    Set Wrk = XLApp.Workbooks.Open(txtFile(Index).Text, False, True)
    'Read all worksheets in xls file
    For Each Sht In Wrk.Worksheets
        'Put the name of worksheet in combo
        cmbSheet(Index).AddItem Sht.Name
    Next
    cmbSheet(Index).ListIndex = 0
    DoEvents
    'Close the XLS file and dont save
    Wrk.Close False
    'Quit the MS Excel
    XLApp.Quit
    'Release variables
    Set XLApp = Nothing
    Set Wrk = Nothing
    Set Sht = Nothing
Else
    'MsgBox "Cancel was pressed"
End If
End Sub

Sub cmdLoad_Click(Index As Integer)
On Error GoTo Oops
Dim c As Integer
Dim XLApp As New Excel.Application
Dim Wrk As Excel.Workbook
Dim Sht As Excel.Worksheet
Dim Rng As Excel.Range
Dim ArrayCells() As Variant
Screen.MousePointer = 11
If cmbSheet(Index).ListIndex <> -1 Then
    'Create a new instance of Excel
    Set XLApp = CreateObject("Excel.Application")
    'Open the XLS file. The two parameters representes, UpdateLink = False and ReadOnly = True. These parameters have this setting to dont occur any error on broken links and allready opened XLS file.
    Set Wrk = XLApp.Workbooks.Open(txtFile(Index).Text, False, True)
    'Set the SHT variable to selected worksheet
    Set Sht = Wrk.Worksheets(cmbSheet(Index).List(cmbSheet(Index).ListIndex))
    'Get the used range of current worksheet
    Set Rng = Sht.UsedRange
    'Change the dimensions of array to fit the used range of worksheet
    ReDim ArrayCells(1 To Rng.Rows.Count, 1 To Rng.Columns.Count)
    'Transfer values of the used range to new array
    If Option1.Value Then
        ArrayCells = Rng.Value
    ElseIf Option2.Value Then
        ArrayCells = Rng.Formula
    End If
    'Close worksheet
    Wrk.Close False
    'Quit the MS Excel
    XLApp.Quit
    'Release variables
    Set XLApp = Nothing
    Set Wrk = Nothing
    Set Sht = Nothing
    Set Rng = Nothing
    'Configure the flexgrid to display data
    With Grid1(Index)
        'we set the redraw to false and make the grid invisible
        'this allows the data to be filled in much faster since the screen
        'does not have to redraw with each change in cell text
        .Redraw = False
        .Visible = False
        'empty the grid
        .Clear
        'set the grid to a few cells to start from scratch for each new file
        .Rows = 1
        .Cols = 1
        .FixedCols = 0
        .FixedRows = 0
        'set the number of rows to selected sheet's rows
        .Rows = UBound(ArrayCells, 1)
        'do the same for the columns
        .Cols = UBound(ArrayCells, 2)
        For r = 0 To UBound(ArrayCells, 1) - 1
            For c = 0 To UBound(ArrayCells, 2) - 1
                'set the text to the cell
                .TextMatrix(r, c) = CStr(ArrayCells(r + 1, c + 1))
            Next
        Next
        .Redraw = True
        .Visible = True
    End With
Else
    MsgBox "Select the worksheet!", vbCritical
    cmbSheet(Index).SetFocus
End If
GoTo Exit_cmdLoad_Click
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine cmdLoad_Click "
EMess$ = "Error # " & err.Number & " - " & err.Description & vbCrLf
EMess$ = EMess$ & "Occurred in cmdLoad_Click"
EMess$ = EMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(EMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_cmdLoad_Click:
Grid1(Index).Redraw = True
Grid1(Index).Visible = True
Refresh
DoEvents
Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
Inits = 0
WindowState = vbMaximized
Show
InitDir$ = GetSetting(App.EXEName, "Settings", "InitDir", "")
cmdOpen_Click (0)
cmdOpen_Click (1)
Inits = 1
Compare
End Sub

Private Sub Form_Resize()
If WindowState = vbMinimized Then Exit Sub
For i = 0 To 1
    With Me.Grid1(i)
        'size the grids to be 1/2 of the form
        .Height = Me.Height - .Top - 550
        .Width = (Me.Width / 2) - 150
        txtFile(i).Width = (.Width - cmdOpen(i).Width) - 100
    End With
Next i
Grid1(0).Left = 50
Grid1(1).Left = Grid1(0).Left + Grid1(0).Width + 50
'size the other objects
For i = 0 To 1
    txtFile(i).Left = Grid1(i).Left
    cmdOpen(i).Left = txtFile(i).Left + txtFile(i).Width + 50
Next i
Text1(1).Left = Grid1(1).Left
Text1(1).Width = Text1(0).Width
cmbSheet(1).Left = Grid1(1).Left
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmResults
End Sub

Private Sub Grid1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim CellData$
With Grid1(Index)
    If .MouseRow >= .Rows Then Exit Sub
    If .MouseCol >= .Cols Then Exit Sub
    If .MouseRow < 1 Then Exit Sub
    If .MouseCol < 1 Then Exit Sub
    'setup the cell's text to display as the tooldtip for the grid
    CellData$ = .TextMatrix(.MouseRow, .MouseCol)
    .ToolTipText = CellData$
End With
End Sub

Private Sub Grid1_RowColChange(Index As Integer)
If Inits = 0 Then Exit Sub
Caption = "Row " & Grid1(Index).Row & " col " & Grid1(Index).Col & " (" & Chr$(65 + Grid1(Index).Col) & (Grid1(Index).Row + 1) & ")"
End Sub

Private Sub Grid1_Scroll(Index As Integer)
If Inits = 0 Then Exit Sub
'if the a grid scrolls, change the opposite one as well
Grid1(1 - Index).Row = Grid1(Index).Row
Grid1(1 - Index).Col = Grid1(Index).Col
Grid1(1 - Index).TopRow = Grid1(Index).TopRow
Grid1(1 - Index).LeftCol = Grid1(Index).LeftCol
End Sub

Private Sub Grid1_SelChange(Index As Integer)
If Inits = 0 Then Exit Sub
'if the selection in a grid changes, change the opposite one as well
Grid1(1 - Index).Row = Grid1(Index).Row
Grid1(1 - Index).Col = Grid1(Index).Col
Grid1(1 - Index).RowSel = Grid1(Index).Row
Grid1(1 - Index).ColSel = Grid1(Index).Col
'display the contents of the selected cell in our textbox
Text1(0) = Grid1(0).Text
Text1(1) = Grid1(1).Text
Refresh
End Sub

Sub Compare()
On Error GoTo Oops
Dim CompDiff As Byte
Dim Comp$(2)
Screen.MousePointer = 11
Inits = 0
Grid1(0).Redraw = False
Grid1(1).Redraw = False
'reset both grids to default colors and fonts
For i = 0 To 1
    'select all of the cells in the grid control
    Grid1(i).Row = 0
    Grid1(i).Col = 0
    Grid1(i).RowSel = Grid1(i).Rows - 1
    Grid1(i).ColSel = Grid1(i).Cols - 1
    'set the fill style to repeat
    Grid1(i).FillStyle = flexFillRepeat
    'set the foreground color to black and background to white
    Grid1(i).CellForeColor = &H80000008
    Grid1(i).CellBackColor = &H80000005
    Grid1(i).CellFontBold = False
    'reset the fillstyle back to single
    Grid1(i).FillStyle = flexFillSingle
    'reset the selection back to one cell
    Grid1(i).Row = 1
    Grid1(i).Col = 1
    Grid1(i).RowSel = 1
    Grid1(i).ColSel = 1
Next i
Dim gRows(2) As Integer
Dim gCols(2) As Integer
gRows(0) = Grid1(0).Rows
gRows(1) = Grid1(1).Rows
gCols(0) = Grid1(0).Cols
gCols(1) = Grid1(1).Cols
'check if the rows and columns are the same in each spreadsheet.
'if they are different, give the user the option to continue or cancel.
If gRows(0) <> gRows(1) Then
    EMess$ = "El archivo de la izquierda tiene " & gRows(0) & " rows."
    EMess$ = EMess$ & vbCrLf & "El archivo de la derecha tiene " & gRows(1) & " rows."
    EMess$ = EMess$ & vbCrLf & "Desea comparar de todas maneras?"
    lRet = MsgBox(EMess$, vbOKCancel + vbCritical, "Spreadsheet Size Difference!")
    If lRet = vbCancel Then GoTo Exit_Compare
    'set the rows to be the same in each spreadsheet
    If gRows(0) > gRows(1) Then
        Grid1(1).Rows = Grid1(0).Rows
    ElseIf gRows(1) > gRows(0) Then
        Grid1(0).Rows = Grid1(1).Rows
    End If
End If
If gCols(0) <> gCols(1) Then
    EMess$ = "El archivo de la izquierda tiene " & gCols(0) & " Cols."
    EMess$ = EMess$ & vbCrLf & "El archivo de la derecha tiene " & gCols(1) & " Cols."
    EMess$ = EMess$ & vbCrLf & "Desea comparar de todas maneras?"
    lRet = MsgBox(EMess$, vbOKCancel + vbCritical, "Spreadsheet Size Difference!")
    If lRet = vbCancel Then GoTo Exit_Compare
    'set the cols to be the same in each spreadsheet
    If gCols(0) > gCols(1) Then
        Grid1(1).Cols = Grid1(0).Cols
    ElseIf gCols(1) > gCols(0) Then
        Grid1(0).Cols = Grid1(1).Cols
    End If
End If
For r = 0 To gRows(0) - 1
    For c = 0 To gCols(0) - 1
        'if we ignore case, set both text's to be upper case
        If chkCase.Value = vbChecked Then
            Comp$(0) = UCase(Grid1(0).TextMatrix(r, c))
            Comp$(1) = UCase(Grid1(1).TextMatrix(r, c))
        Else
            Comp$(0) = Grid1(0).TextMatrix(r, c)
            Comp$(1) = Grid1(1).TextMatrix(r, c)
        End If
        'compare the contents of the cells in the 2 grids
        If Comp$(0) <> Comp$(1) Then
            'we found a difference
            For i = 0 To 1
                'select the cell that's difference
                Grid1(i).Row = r
                Grid1(i).Col = c
                'set the color to white on red
                Grid1(i).CellForeColor = vbWhite
                Grid1(i).CellBackColor = vbRed
                'set the font to bold
                Grid1(i).CellFontBold = True
                'Grid1(i).TopRow = r
                'Grid1(i).LeftCol = c
                CompDiff = 1
            Next i
            'add the mismatched data to the results sheet
            frmResults.GridResults.AddItem r & vbTab & c & vbTab & Grid1(0).TextMatrix(r, c) & vbTab & Grid1(1).TextMatrix(r, c)
        End If
    Next c
Next r
If frmResults.GridResults.Rows = 1 Then
    'if we haven't added anything, unload the results form
    Unload frmResults
Else
    frmResults.Left = (Me.Width - frmResults.Width) / 2
    frmResults.Height = ((frmResults.GridResults.Rows + 2) * (frmResults.GridResults.RowHeight(1))) + 200
    frmResults.Top = (Me.Height - frmResults.Height) - 150
End If
'show a message representing the compare status
If CompDiff = 1 Then
    EMess$ = "Ambos archivos son diferentes!"
    'MsgBox EMess$, vbCritical
    lblResults.ForeColor = vbWhite
    lblResults.BackColor = vbRed
Else
    EMess$ = "Excel Files are the same!"
    lblResults.ForeColor = 0
    lblResults.BackColor = Me.BackColor
End If
lblResults.Caption = EMess$
GoTo Exit_Compare
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine Compare "
EMess$ = "Error # " & err.Number & " - " & err.Description & vbCrLf
EMess$ = EMess$ & "Occurred in Compare"
EMess$ = EMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(EMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_Compare:
Grid1(0).Redraw = True
Grid1(1).Redraw = True
Inits = 1
Screen.MousePointer = 0
End Sub
