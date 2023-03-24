VERSION 5.00
Begin VB.Form FrmPass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Autenticacion"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   5580
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   405
      Left            =   2070
      TabIndex        =   2
      Top             =   1110
      Width           =   1425
   End
   Begin VB.TextBox TxtPass 
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
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   870
      PasswordChar    =   "?"
      TabIndex        =   1
      Text            =   "password"
      Top             =   570
      Width           =   3855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "INGRESE CLAVE DE ACCESO"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "FrmPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If Trim(UCase(TxtPass.Text)) = "170976" Then
    FrmInit.Show
    Unload Me
Else
    MsgBox "Autenticación incorrecta!"
    End
End If

End Sub

Private Sub TxtPass_Click()

TxtPass.SelStart = 0
TxtPass.SelLength = Len(TxtPass.Text)

End Sub

Private Sub TxtPass_GotFocus()

TxtPass.SelStart = 0
TxtPass.SelLength = Len(TxtPass.Text)

End Sub

Private Sub TxtPass_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    Command1.SetFocus
End If

End Sub
