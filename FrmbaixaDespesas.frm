VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FrmBaixaDespesas 
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
   Begin VB.TextBox Codigo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   6000
      TabIndex        =   19
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Busca Desp F4"
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
      Left            =   8520
      TabIndex        =   16
      Top             =   1080
      Width           =   2055
   End
   Begin MSMask.MaskEdBox Mask 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2880
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yy"
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
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
      TabIndex        =   18
      Top             =   120
      Width           =   7815
   End
   Begin VB.CommandButton CmdFechar 
      Caption         =   "&Fechar F10"
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
      Left            =   8520
      TabIndex        =   17
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar F3"
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
      Left            =   8520
      TabIndex        =   15
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok F2"
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
      Left            =   8520
      TabIndex        =   14
      Top             =   120
      Width           =   2055
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
      TabStop         =   0   'False
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
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   1935
   End
   Begin MSMask.MaskEdBox Mask 
      Height          =   375
      Index           =   1
      Left            =   3120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2880
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Mask 
      Height          =   375
      Index           =   2
      Left            =   2160
      TabIndex        =   1
      Top             =   1080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yy"
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Mask 
      Height          =   375
      Index           =   3
      Left            =   6240
      TabIndex        =   3
      Top             =   1080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Mask 
      Height          =   375
      Index           =   4
      Left            =   4200
      TabIndex        =   2
      Top             =   1080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   " "
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
      BackColor       =   &H00D8C5B6&
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
      BackColor       =   &H00D8C5B6&
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
      BackColor       =   &H00D8C5B6&
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
      TabIndex        =   11
      Top             =   1680
      Width           =   1395
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00D8C5B6&
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
      Left            =   6240
      TabIndex        =   10
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00D8C5B6&
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
      TabIndex        =   9
      Top             =   720
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00D8C5B6&
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
      TabIndex        =   8
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00D8C5B6&
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
      TabIndex        =   7
      Top             =   720
      Width           =   570
   End
End
Attribute VB_Name = "FrmBaixaDespesas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RsDespesa As Recordset
Private a As Integer
Private Sub CmdCancelar_Click()
LimpaControle
End Sub

Private Sub CmdCancelar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{B}"
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub CmdFechar_Click()
On Error Resume Next

Unload Me
End Sub

Private Sub CmdFechar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{B}"
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub CmdOk_Click()
SalvaAlteracao
LimpaControle
End Sub

Private Sub CmdOk_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{B}"
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Command1_Click()
Despesasnaoquitadas.Show , Me
End Sub

Private Sub Form_Activate()
Set GlFormA = Me
End Sub

Private Sub Form_Load()
On Error Resume Next
Titulo.Text = "Aguardando Documento"
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
Set Dbbase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
Set RsDespesa = Dbbase.OpenRecordset("ALID014", dbOpenDynaset, dbSeeChanges, dbOptimistic)
'Txt(2).Text = 0
'Txt(6).Text = 0

End Sub
Function buscaDoc(Optional CodigoTxt As String = "")
On Error GoTo erroBusca
Dim RsFornecedor As Recordset
Dim a As Long
Dim LcCriterio As String
If Len(Trim(txt(0).Text)) = 0 And Len(codigo.Text) = 0 Then Exit Function
AbreBase
Set RsFornecedor = Dbbase.OpenRecordset("ALID002", dbOpenDynaset, dbSeeChanges, dbOptimistic)

If Len(CodigoTxt) > 0 Then
   LcCriterio = "codigo=" & CodigoTxt
Else
    If Len(txt(0).Text) > 0 Then
       LcCriterio = "nf='" & txt(0).Text & "'"
    End If
End If

RsDespesa.FindFirst LcCriterio
If Not RsDespesa.NoMatch Then
  If IsNull(RsDespesa!credor) Then
        txt(4).Text = ""
  Else
        LcCriterioCli = "Codigo='" & RsDespesa!credor & "'"
        RsFornecedor.FindFirst LcCriterioCli
        If Not RsFornecedor.NoMatch Then
           txt(4).Text = RsFornecedor!RazaoSoc
        Else
           txt(4).Text = ""
        End If
  End If
  If Not IsNull(RsDespesa!DTVENC) Then Mask(0).Text = Format(RsDespesa!DTVENC, "dd/mm/yy") Else Mask(0).Text = "  /  /  "
  If Not IsNull(RsDespesa!Valor) Then Mask(1).Text = RsDespesa!Valor Else Mask(1).Text = 0
  If Not IsNull(RsDespesa!Valor) Then Mask(3).Text = RsDespesa!Valor Else Mask(3).Text = 0
  
  cmdOK.Enabled = True
  If RsDespesa!DTPAGTO <> "" Then
      If Not IsNull(RsDespesa!DTPAGTO) Then Mask(2).Text = Format(RsDespesa!DTPAGTO, "dd/mm/yy") Else Mask(2).Text = "  /  /  "
      If Not IsNull(RsDespesa!Acrescimo) Then Mask(4).Text = RsDespesa!Acrescimo Else Mask(4).Text = 0
      If Not IsNull(RsDespesa!Valor) Then Mask(3).Text = RsDespesa!Valor Else Mask(3).Text = 0

     Titulo.Text = "Documento Quitado"
     cmdOK.Enabled = False
  Else
     Mask(2).Text = "  /  /  "
     Mask(4).Text = ""
     Titulo.Text = "Documento Aberto"
  End If
Else
  MsgBox "Documento Não Encontrado...", 48, "Aviso"
  LimpaControle
End If
RsFornecedor.Close
Set RsFornecedor = Nothing

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
Set RsDespesa = Nothing
Set Dbbase = Nothing
FrmPrincipal.SetFocus
End Sub

Private Sub Mask_Change(Index As Integer)
If Index = 4 Then CalculoAcresino
End Sub

Private Sub Mask_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{B}"
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Mask_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 3 Then If KeyAscii = 46 Then KeyAscii = 44
If Index = 4 Then If KeyAscii = 46 Then KeyAscii = 44
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{B}"
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Txt_LostFocus(Index As Integer)
Select Case Index
    Case Is = 0
       If Len(codigo.Text) > 0 Then
          buscaDoc codigo.Text
       Else
          buscaDoc
       End If
       
       
End Select

End Sub
Function SalvaAlteracao()
On Error GoTo errosalva
AbreBase
Dim a As Long

Dim LcCriterio, LcTipoMOne As String
Dim LcValor As Double

VALPAGO = Replace(CCur(Mask(3).Text), ",", ".")


LCSqlInc = "Update alid014 SET "
LCSqlInc = LCSqlInc & "dtpagto=#" & Format(Mask(2).Text, "mm/dd/yy") & "#,"
If Len(Mask(4).Text) = 0 Then
   LCSqlInc = LCSqlInc & "acrescimo=0,"
Else
  LCSqlInc = LCSqlInc & "acrescimo=" & Replace(CCur(Mask(4).Text), ",", ".") & ","
End If

LCSqlInc = LCSqlInc & "ValPago=" & VALPAGO
If Len(codigo.Text) > 0 Then
   LCSqlInc = LCSqlInc & " where  codigo=" & codigo.Text
Else
   LCSqlInc = LCSqlInc & " where  NF='" & txt(0).Text & "'"
End If
 Dbbase.Execute LCSqlInc
LcComentario = "-Incluido Nota Fiscal- Efetuando a Inclusão na Tabela."
'LcRegistrosAfetados = ExecutaSql(LCSqlInc)

'RsDespesa.Edit
'RsDespesa!DTPAGTO = Mask(2).Text
'If Len(Mask(4).Text) = 0 Then RsDespesa!acrescimo = 0 Else RsDespesa!acrescimo = CCur(Mask(4).Text)
'RsDespesa!VALPAGO = CCur(Mask(3).Text)
'RsDespesa!quitado = True

'RsDespesa.Update
LcTipoMOne = RsDespesa!TPMONET & ""
LcValor = CCur(Mask(3).Text)
If GlBaixaDespesa Then Call lancacaixa("Despesas", txt(0).Text, LcTipoMOne, LcValor)

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
    txt(a).Text = ""
    
Next
Mask(0).Text = "  /  /  "
Mask(1).Text = ""
Mask(2).Text = "  /  /  "
Mask(3).Text = ""
Mask(4).Text = ""
codigo.Text = ""
txt(0).SetFocus
End Function
Function CalculoAcresino()
On Error Resume Next
Dim LcDevido, LcACrescimo, LcTotal As Currency
If IsNull(Mask(1).Text) Then LcDevido = 0 Else LcDevido = CCur(Mask(1).Text)
If IsNull(Mask(4).Text) Then LcACrescimo = 0 Else LcACrescimo = CCur(Mask(4).Text)
LcTotal = LcDevido + LcACrescimo
Mask(3).Text = LcTotal
End Function
