VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form frmResults 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Comparar Resultados"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   8430
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmResults.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   8430
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid GridResults 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   9975
      _Version        =   393216
      Cols            =   4
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
End
Attribute VB_Name = "frmResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gFormat$

Private Sub Form_Load()
gFormat$ = "<Row |<Col  |<-             File1             |<-             File2             "
With frmResults.GridResults
    .Clear
    .Rows = 1
    '.TextMatrix(0, 0) = "Row"
    '.TextMatrix(0, 1) = "Col"
    '.TextMatrix(0, 2) = "File1"
    '.TextMatrix(0, 3) = "File2"
    'formatstring does NOT work except with fixed rows!!!
    .FormatString = gFormat$
End With
Me.ZOrder 0
Show
End Sub

Private Sub Form_Resize()
Dim cWidth
If Me.Height > Screen.Height Then
    Me.Height = Screen.Height - 100
    Me.Top = 50
End If
With GridResults
    .Left = 50
    .Top = 50
    .Width = Me.Width - 200
    .Height = Me.Height - 450
    cWidth = .ColWidth(0) + .ColWidth(1)
    .ColWidth(2) = (Me.Width / 2) - (cWidth - 100)
    .ColWidth(3) = (Me.Width / 2) - (cWidth - 100)
End With
End Sub

Private Sub GridResults_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim CellData$
With GridResults
    If .MouseRow >= .Rows Then Exit Sub
    If .MouseCol >= .Cols Then Exit Sub
    If .MouseRow < 1 Then Exit Sub
    If .MouseCol < 1 Then Exit Sub
    CellData$ = "Row " & .MouseRow & "/Col " & .MouseCol & " File 1=>" & .TextMatrix(.MouseRow, 2)
    CellData$ = CellData$ & " - File 2=>" & .TextMatrix(.MouseRow, 3)
    .ToolTipText = CellData$
End With
End Sub
