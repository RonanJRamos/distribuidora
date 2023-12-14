VERSION 5.00
Begin VB.Form FrmBaixaDespesa 
   BackColor       =   &H00D8C5B6&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Baixa em Despesas"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10710
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   10710
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Titulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   240
      TabIndex        =   17
      Top             =   120
      Width           =   7815
   End
   Begin VB.CommandButton CmdFechar 
      Caption         =   "&Fechar"
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
      Left            =   8640
      TabIndex        =   16
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
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
      Left            =   8640
      TabIndex        =   15
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
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
      Left            =   8640
      TabIndex        =   14
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox Txt 
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
      Index           =   6
      Left            =   3120
      TabIndex        =   11
      Text            =   "0"
      Top             =   2880
      Width           =   2295
   End
   Begin VB.TextBox Txt 
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
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Top             =   2880
      Width           =   2295
   End
   Begin VB.TextBox Txt 
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
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   7815
   End
   Begin VB.TextBox Txt 
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
      Index           =   3
      Left            =   6120
      TabIndex        =   3
      Text            =   "0"
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox Txt 
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
      Index           =   2
      Left            =   4200
      TabIndex        =   2
      Text            =   "0"
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox Txt 
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
      Index           =   1
      Left            =   2160
      TabIndex        =   1
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox Txt 
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
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   495
      Left            =   8520
      TabIndex        =   18
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000001&
      BorderWidth     =   2
      X1              =   8400
      X2              =   8400
      Y1              =   0
      Y2              =   3360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      BorderWidth     =   2
      X1              =   0
      X2              =   8400
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Doc."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   6
      Left            =   3120
      TabIndex        =   13
      Top             =   2520
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   5
      Left            =   120
      TabIndex        =   12
      Top             =   2520
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fornecedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   4
      Left            =   240
      TabIndex        =   9
      Top             =   1680
      Width           =   1395
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Pago"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   6120
      TabIndex        =   8
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Acrescímo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   4200
      TabIndex        =   7
      Top             =   720
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Pag."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   2160
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Doc."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   570
   End
End
Attribute VB_Name = "FrmBaixaDespesa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RsDespesa As Recordset
Private Sub CmdCancelar_Click()
LimpaControle
End Sub

Private Sub CmdFechar_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub CmdOk_Click()
SalvaAlteracao
LimpaControle
End Sub

Private Sub Form_Load()
Titulo.Text = "Aguardando Documento"
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
Set Dbbase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
Set RsDespesa = Dbbase.OpenRecordset("ContasPagar", dbOpenDynaset)
Txt(2).Text = 0
Txt(6).Text = 0

End Sub
Function buscaDoc()
On Error GoTo erroBusca

Dim a As Long

Dim LcCriterio As String


LcCriterio = "Documento='" & Txt(0).Text & "'"
RsDespesa.FindFirst LcCriterio
If Not RsDespesa.NoMatch Then
  If IsNull(RsDespesa!fornecedor) Then Txt(4).Text = "" Else Txt(4).Text = RsDespesa!fornecedor
  Txt(5).Text = RsDespesa!Vencimento
  Txt(6).Text = Format(RsDespesa!valor, "Currency")
  Txt(3).Text = Format(RsDespesa!valor, "Currency")
  cmdOK.Enabled = True
  If RsDespesa!quitado Then
     Txt(1).Text = RsDespesa!DataPag
     Txt(2).Text = Format(RsDespesa!acrescimo, "Currency")
     Txt(3).Text = RsDespesa!ValorPago
     Titulo.Text = "Documento Quitado"
     cmdOK.Enabled = False
  Else
     Txt(1).Text = ""
     Txt(2).Text = 0
     Txt(3).Text = 0
     Titulo.Text = "Documento Aberto"
  End If
Else
  MsgBox "Documento Não Encontrado...", 48, "Aviso"
  LimpaControle
End If


Exit Function
erroBusca:
Select Case ErrosSistema
       Case Is = 0
          Resume Next
       Case Is = 6
          Resume 0
       Case Is = 7
          End
       Case Else
         If err = 13 Then
             MsgBox "Preenchimento de campo Incorreto. Favor Verificar.", 48, "Aviso"
             Exit Function
         Else
             MsgBox err.Description & " N " & err
             Resume Next
         End If
End Select


End Function

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
RsDespesa.Close
Dbbase.Close

End Sub

Private Sub Txt_Change(Index As Integer)
If Index = 2 Then CalculoAcresino
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   Select Case Index
        Case Is = 0
             buscaDoc
             Txt(1).SetFocus
        Case Is = 1
             Txt(2).SetFocus
        Case Is = 2
             Txt(3).SetFocus
        Case Is = 3
             cmdOK.SetFocus
   End Select
End If
End Sub

Private Sub Txt_LostFocus(Index As Integer)
Select Case Index
    Case Is = 0
       buscaDoc
       
End Select

End Sub
Function SalvaAlteracao()
On Error GoTo errosalva

Dim a As Long

Dim LcCriterio As String
Dim LcValor As Currency
RsDespesa.Edit
RsDespesa!DataPag = Format(Txt(1).Text, "dd/mm/yyyy")
RsDespesa!acrescimo = CCur(Txt(2).Text)
RsDespesa!ValorPago = CCur(Txt(3).Text)
RsDespesa!quitado = True

RsDespesa.Update
LcValor = CCur(Txt(3).Text)
lancacaixaDebito (LcValor)

Exit Function
errosalva:
Select Case ErrosSistema
       Case Is = 0
          Resume Next
       Case Is = 6
          Resume 0
       Case Is = 7
          End
       Case Else
         If err = 13 Then
             MsgBox "Preenchimento de campo Incorreto. Favor Verificar.", 48, "Aviso"
             Exit Function
         Else
             MsgBox err.Description & " N " & err
             Resume Next
         End If
End Select

End Function
Function LimpaControle()
On Error Resume Next
Dim a As Integer
For a = 0 To 6
    Txt(a).Text = ""
Next
Txt(0).SetFocus
End Function
Function CalculoAcresino()
On Error Resume Next
Dim LcDevido, LcACrescimo, LcTotal As Currency
If IsNull(Txt(6).Text) Then LcDevido = 0 Else LcDevido = CCur(Txt(6).Text)
If IsNull(Txt(2).Text) Then LcACrescimo = 0 Else LcACrescimo = CCur(Txt(2).Text)
LcTotal = LcDevido + LcACrescimo
Txt(3) = Format(LcTotal, "Currency")

End Function
