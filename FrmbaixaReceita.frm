VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FrmBaixaReceita 
   BackColor       =   &H00E4E3D6&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Baixa em Receitas"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10785
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   10785
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox TipoBaixa 
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
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2160
      Width           =   8055
   End
   Begin VB.TextBox Codigo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   8640
      TabIndex        =   21
      Top             =   2520
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox CodCliente 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   8640
      TabIndex        =   20
      Top             =   2040
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Busca Rec F4"
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
      TabIndex        =   17
      Top             =   1080
      Width           =   1935
   End
   Begin MSMask.MaskEdBox Mask 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4080
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
      BackColor       =   &H80000004&
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
      TabIndex        =   19
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
      Left            =   8640
      TabIndex        =   18
      Top             =   1560
      Width           =   1935
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
      Left            =   8640
      TabIndex        =   16
      Top             =   600
      Width           =   1935
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
      Left            =   8640
      TabIndex        =   15
      Top             =   120
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
      Index           =   4
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3240
      Width           =   8055
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
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4080
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
   Begin VB.Label Label2 
      BackColor       =   &H00E4E3D6&
      Caption         =   "Motivo da Baixa"
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
      Left            =   120
      TabIndex        =   22
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000001&
      BorderWidth     =   2
      X1              =   8400
      X2              =   8400
      Y1              =   0
      Y2              =   4680
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
      TabIndex        =   14
      Top             =   3720
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
      TabIndex        =   13
      Top             =   3720
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
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
      TabIndex        =   12
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Recebido"
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
      TabIndex        =   11
      Top             =   720
      Width           =   1860
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
      TabIndex        =   10
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
      TabIndex        =   9
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
      TabIndex        =   8
      Top             =   720
      Width           =   570
   End
End
Attribute VB_Name = "FrmBaixaReceita"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private a As Long
Private LCSqlInc As String
Private RsReceita As ADODB.Recordset
Private Sub CmdCancelar_Click()
LimpaControle
End Sub

Private Sub CmdCancelar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
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
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub CmdOk_Click()
Dim LcPodeBaixar As Boolean
LcPodeBaixar = False

If Not IsNumeric(Mask(3).Text) Then
    MsgBox "Informe o valor do pagamento", 64, "Aviso"
Else
   If CLng(Mask(3).Text) <= 0 Then
      MsgBox "Informe o valor do pagamento", 64, "Aviso"
   Else
      LcPodeBaixar = True
   End If
End If
If LcPodeBaixar Then
    If Not IsDate(Mask(2).Text) Then
       LcPodeBaixar = False
       MsgBox "Informe a data de pagamento", 64, "Aviso"
    Else
        LcPodeBaixar = True
    End If
End If
If LcPodeBaixar Then
    SalvaAlteracao
    LimpaControle
End If
End Sub

Private Sub CmdOk_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 115 Then SendKeys "%+{B}"
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Command1_Click()
On Error Resume Next
receitasnaoquitadas.Show , Me
Txt(0).SetFocus
End Sub

Private Sub Form_Activate()
Set GlFormA = Me
End Sub

Private Sub Form_Load()

Titulo.Text = "Aguardando Documento"
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
CarregaTipoBaixa
'Set Dbbase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
'Set RsReceita = Dbbase.OpenRecordset("ALID015", dbOpenDynaset)
'Txt(2).Text = 0
'Txt(6).Text = 0
'abreconexao
End Sub
Function buscaDoc()
On Error GoTo erroBusca
Dim rsCliente As Recordset
Dim a As Long
Dim LcCriterio As String
If Len(Trim(Txt(0).Text)) = 0 Then Exit Function
AbreBase
Set rsCliente = Dbbase.OpenRecordset("ALID001", dbOpenDynaset, dbSeeChanges, dbOptimistic)
If IsNumeric(codigo.Text) Then
   If codigo.Text > 0 Then
      LcCriterio = "select * from alid015 where Codigo=" & codigo.Text
   Else
      LcCriterio = "select * from alid015 where NF='" & Txt(0).Text & "'"
   End If
Else
  LcCriterio = "select * from alid015 where NF='" & Txt(0).Text & "'"
End If

Set RsReceita = AbreRecordset(LcCriterio)

If Not RsReceita.EOF Then
  If IsNull(RsReceita!Cliente) Then
        Txt(4).Text = ""
  Else
        LcCriterioCli = "Codigo='" & RsReceita!Cliente & "'"
        rsCliente.FindFirst LcCriterioCli
        If Not rsCliente.NoMatch Then
           Txt(4).Text = rsCliente!RAZAOSOC
           CodCliente.Text = rsCliente!codigo
        Else
           Txt(4).Text = ""
        End If
  End If
  If Not IsNull(RsReceita!DTVENC) Then Mask(0).Text = RsReceita!DTVENC Else Mask(0).Text = "  /  /  "
  If Not IsNull(RsReceita!Valor) Then Mask(1).Text = RsReceita!Valor Else Mask(1).Text = 0
  If Not IsNull(RsReceita!Valor) Then Mask(3).Text = RsReceita!Valor Else Mask(3).Text = 0
  
  cmdOK.Enabled = True
  If RsReceita!DTPAGTO <> "" Then
      If Not IsNull(RsReceita!DTPAGTO) Then Mask(2).Text = RsReceita!DTPAGTO Else Mask(2).Text = "  /  /  "
      If Not IsNull(RsReceita!Acrescimo) Then Mask(4).Text = RsReceita!Acrescimo Else Mask(4).Text = 0
      If Not IsNull(RsReceita!Valor) Then Mask(3).Text = RsReceita!Valor Else Mask(3).Text = 0

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
'RsCliente.Close
'Set RsCliente = Nothing

Exit Function
erroBusca:
Select Case ErrosSistema
       Case Is = 0
         ' Resume 0
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
RsReceita.Close
Dbbase.Close
Set RsReceita = Nothing
Set Dbbase = Nothing
FrmPrincipal.SetFocus
'FechaConexao
End Sub

Private Sub Mask_Change(Index As Integer)
If Index = 4 Then CalculoAcresino
End Sub

Private Sub Mask_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 115 Then SendKeys "%+{B}"
If KeyCode = 114 Then SendKeys "%+{C}"
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
       buscaDoc
End Select

End Sub
Function SalvaAlteracao()
On Error GoTo errosalva

Dim a As Long

Dim LcCriterio, LcTipoMOne As String
Dim LcACrescimo As String
Dim rsCliente As Recordset
Dim LcCodigoTipoRec As String

Set rsCliente = Dbbase.OpenRecordset("ALID001", dbOpenDynaset, dbSeeChanges, dbOptimistic)
If Len(Mask(4).Text) = 0 Then LcACrescimo = 0 Else LcACrescimo = CCur(Mask(4).Text)

'LcValor = CCur(Mask(4).Text)
LcACrescimo = Replace(LcACrescimo, ",", ".")

If Len(Mask(3).Text) = 0 Then VALPAGO = 0 Else VALPAGO = CCur(Mask(3).Text)
If Len(TipoBaixa.Text) > 0 Then LcCodigoTipoRec = BuscaCodigoTipoRecebimento(TipoBaixa.Text)
'LcValor = CCur(Mask(4).Text)
VALPAGO = Replace(VALPAGO, ",", ".")


LCSqlInc = "Update alid015 SET "
LCSqlInc = LCSqlInc & "dtpagto='" & Format(Mask(2).Text, "yy-mm-dd") & "',"
LCSqlInc = LCSqlInc & "ValPago=" & VALPAGO & ","
LCSqlInc = LCSqlInc & "Acrescimo=" & LcACrescimo & ","
LCSqlInc = LCSqlInc & "codDesp='" & LcCodigoTipoRec & "',"
LCSqlInc = LCSqlInc & "NomeDesp='" & TipoBaixa.Text & "'"
LCSqlInc = LCSqlInc & " where  NF='" & Txt(0).Text & "'"
 Debug.Print LCSqlInc
LcComentario = "-Incluido Nota Fiscal- Efetuando a Inclusão na Tabela."
LcRegistrosAfetados = ExecutaSql(LCSqlInc)
Debug.Print conexaoAdo.ConnectionString



'RsReceita.Edit
'RsReceita!DTPAGTO = Mask(2).Text
'RsReceita!VALPAGO = CCur(Mask(3).Text)
'RsReceita!quitado = True
LcCriterio = "CODIGO='" & CodCliente.Text & "'"
'RsReceita.Update

LcValor = CCur(Mask(3).Text)
'Call lancacaixa("Receita", Txt(0).Text,)
rsCliente.FindFirst LcCriterio
If Not rsCliente.NoMatch Then
   Dim LCValorUtil As Double
   LCValorUtil = rsCliente("CreditoUtilizado") - LcValor
   If LCValorUtil < 0 Then LCValorUtil = 0
   rsCliente.Edit
   rsCliente("CreditoUtilizado") = LCValorUtil ' rsCliente("CreditoUtilizado") - LcValor
   rsCliente.Update
End If
If GlBaixaReceita Then
   LcTipoMOne = RsReceita!TPMONET & ""
   'Call lancacaixa("Receita", Txt(0).Text, LcTipoMOne, LcValor)
End If
rsCliente.Close
Set rsCliente = Nothing

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
Sub CarregaTipoBaixa()
Dim RsDesp As Recordset
AbreBase
Set RsDesp = Dbbase.OpenRecordset("select * from alid007 where RD='R'")
 TipoBaixa.AddItem ("")
Do Until RsDesp.EOF
   TipoBaixa.AddItem (RsDesp!Nome & "")
   RsDesp.MoveNext
Loop
   
End Sub
Function BuscaCodigoTipoRecebimento(Nome As String) As String

Dim Resposta As String
Dim RsDesp As Recordset
AbreBase
Set RsDesp = Dbbase.OpenRecordset("select * from alid007 where RD='R' and nome='" & Nome & "'")

If Not RsDesp.EOF Then
   Resposta = RsDesp!cod & ""
  End If
   


BuscaCodigoTipoRecebimento = Resposta
End Function
Function LimpaControle()
On Error Resume Next
Dim a As Long
For a = 0 To 6
    Txt(a).Text = ""
Next
Mask(0).Text = "  /  /  "
Mask(1).Text = ""
Mask(2).Text = "  /  /  "
Mask(3).Text = ""
Mask(4).Text = ""
TipoBaixa.ListIndex = 0
CodCliente.Text = ""
codigo.Text = ""
Txt(0).SetFocus
End Function
Function CalculoAcresino()
On Error Resume Next
Dim LcDevido As Currency, LcACrescimo As Currency, LcTotal As Currency
If IsNull(Mask(1).Text) Then LcDevido = 0 Else LcDevido = CCur(Mask(1).Text)
If IsNull(Mask(4).Text) Then LcACrescimo = 0 Else LcACrescimo = CCur(Mask(4).Text)
LcTotal = LcDevido + LcACrescimo
Mask(3).Text = LcTotal
End Function
