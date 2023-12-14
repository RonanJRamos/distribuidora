VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FrmReajustaPreco 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reajuste de Preços"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9270
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   9270
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox fornecedor 
      Height          =   315
      Left            =   4800
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   240
      Width           =   4215
   End
   Begin VB.TextBox Codigo 
      Height          =   375
      Left            =   7080
      TabIndex        =   16
      Top             =   3000
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CheckBox option1 
      Caption         =   "Guardar Preço Antigo"
      Height          =   195
      Left            =   240
      TabIndex        =   15
      Top             =   2160
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSMask.MaskEdBox txt 
      Height          =   375
      Left            =   4800
      TabIndex        =   6
      Top             =   2160
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   " "
   End
   Begin VB.TextBox bo 
      Enabled         =   0   'False
      Height          =   375
      Left            =   4800
      TabIndex        =   5
      Top             =   960
      Width           =   3615
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de Reajuste"
      Height          =   1695
      Left            =   2400
      TabIndex        =   13
      Top             =   240
      Width           =   2295
      Begin VB.OptionButton Percentual 
         Caption         =   "Percentual"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton Reais 
         Caption         =   "Reais"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Visible         =   0   'False
         Width           =   1815
      End
   End
   Begin VB.CommandButton CmdFechar 
      Caption         =   "&Fechar F10"
      Height          =   495
      Left            =   1920
      TabIndex        =   8
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok F2"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Reajustar"
      Height          =   1695
      Left            =   240
      TabIndex        =   9
      Top             =   240
      Width           =   2055
      Begin VB.OptionButton Arquivo 
         Caption         =   "Arquivo Anterior"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1200
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton Distribuidor 
         Caption         =   "Distribuidor"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   2184
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.OptionButton Individual 
         Caption         =   "Individual"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   780
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton Todas 
         Caption         =   "Todos Produtos"
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fornecedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   4800
      TabIndex        =   19
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Para Reduçao Utilize o Sinal de Menos (-) antes do numero Ex: -10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00649766&
      Height          =   495
      Left            =   4800
      TabIndex        =   17
      Top             =   1440
      Width           =   3615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   6480
      TabIndex        =   14
      Top             =   2400
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Percentual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   4800
      TabIndex        =   11
      Top             =   1920
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Produtos inicados por"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   4800
      TabIndex        =   10
      Top             =   720
      Width           =   2295
   End
End
Attribute VB_Name = "FrmReajustaPreco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type TipoFor
      Codigo As String
      Nome As String
End Type
Dim LcEscolha, a As Long
Dim StrSql As String
Private Mtfor() As TipoFor
Private LcTamanho As Long
Private Sub Arquivo_Click()
On Error Resume Next
LcEscolha = 3
bo.Text = ""
txt.Text = ""
cbo.Enabled = False
Label1(0).Enabled = False
Label1(1).Enabled = False
txt.Enabled = False
option1.Enabled = False
Frame2.Enabled = False
bo.Enabled = False
CmdOk.SetFocus
End Sub

Private Sub Arquivo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{F}"

End Sub

Private Sub bo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then txt.SetFocus
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub bo_LostFocus()
'BuscaProduto
End Sub

Private Sub CmdFechar_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub CmdFechar_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
txt.SetFocus
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{F}"
  End Sub


Private Sub CmdOk_Click()
'On Error Resume Next
Dim LcCriSql As String
Dim LCLEtra As String
Dim LcValor As Currency
Dim RsAtualizacao As ADODB.Recordset
Dim RsAnt As ADODB.Recordset
Dim LcPerAtualizacao As Double
Dim LcCodFornecedor As String
Dim a As Integer
If LcEscolha <> 6 And Len(fornecedor.Text) = 0 Then
  If Len(Trim(txt.Text)) = 0 And LcEscolha <> 3 Then
   MsgBox "É Necessário digitar um Valor para o Reajuste, ou selecionar um fornecedor.", 48, "Aviso"
   txt.SetFocus
   Exit Sub
  End If
End If

If Len(fornecedor.Text) > 0 Then
   For x = 0 To UBound(Mtfor) - 1
       If UCase(Mtfor(x).Nome) = UCase(fornecedor.Text) Then
          LcCodFornecedor = Mtfor(x).Codigo
          Exit For
       End If
   Next
   
End If
If Todas Then
   LcCriSql = "select * From produtos"
   If Len(LcCodFornecedor) > 0 Then LcCriSql = LcCriSql & " where Fornecedor='" & LcCodFornecedor & "'"
End If
If Individual Then
   LcCriSql = "select * From produtos where NOME like '" & UCase(bo.Text) & "%'"
   If Len(LcCodFornecedor) > 0 Then LcCriSql = LcCriSql & " and Fornecedor='" & LcCodFornecedor & "'"
End If
If Arquivo Then
   LcCriSql = "select * From produtos"
    If Len(LcCodFornecedor) > 0 Then LcCriSql = LcCriSql & " where Fornecedor='" & LcCodFornecedor & "'"
End If

'Set Dbbase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
Set RsAtualizacao = AbreRecordset(LcCriSql)

If RsAtualizacao.EOF Then
   MsgBox "Não Exite Produto Com Este Critério...", 48, "Aviso"
   Exit Sub
End If
RsAtualizacao.MoveLast

TotalReg = RsAtualizacao.RecordCount
RsAtualizacao.MoveFirst
LcCapt = Me.Caption
a = 1
   ' Open file for output.

If LcEscolha <> 3 Then
   If option1.Value = 1 Then
      ExecutaSql "delete from precoatuali"
   End If
   Set RsAnt = AbreRecordset("select * from PRECOATUALI")
   Do Until RsAtualizacao.EOF
      Me.Caption = "Atualizando Registro " & a & " de " & TotalReg
    '  If option1 Then Write #Fnum, Left(RsAtualizacao!COD & "          ", 10) & Left(RsAtualizacao!Ptab & "          ", 10) _
      & Left(RsAtualizacao!Lucro & "          ", 10) & Left(RsAtualizacao!MPVENDA & "          ", 10)
      RsAnt.AddNew
      RsAnt!Codigo = RsAtualizacao!Codigo
      RsAnt!Preco = RsAtualizacao!Preco
      RsAnt!Lucro = RsAtualizacao!Lucro
      RsAnt!MiminoVenda = RsAtualizacao!MinimoVenda
      
      If Reais Then

      Else
        If Not IsNull(RsAtualizacao!LimiteVenda) Then
            LcValorLimite = (RsAtualizacao!LimiteVenda * (1 + (CCur(txt.Text) / 100)))
            LcValorLimite = CDbl(AcertaNumero(CStr(LcValorLimite), GlDecimais))
        Else
           LcValorLimite = 0
        End If
        

        LcValorVenda = (RsAtualizacao!Preco * (1 + (CCur(txt.Text) / 100)))
        LcPerAtualizacao = CCur(txt.Text)

        LcValorVenda = CDbl(AcertaNumero(CStr(LcValorVenda), GlDecimais))
        '==> Calcula o minimo
        LcValorMinimo = (RsAtualizacao!MinimoVenda * (1 + (CCur(txt.Text) / 100)))
        LcValorMinimo = CDbl(AcertaNumero(CStr(LcValorMinimo), GlDecimais))
        StrSql = "Update Produtos Set PRECO=" & Replace(LcValorVenda, ",", ".") & "," & _
                 "LimiteVenda=" & Replace(LcValorLimite, ",", ".") & _
                 ",minimoVenda=" & Replace(LcValorMinimo, ",", ".") & _
                 " Where Codigo=" & RsAtualizacao!Codigo
                'Debug.Print StrSql
        Afetado = ExecutaSql(StrSql)
       ' RsAtualizacao.Update
      End If
      If option1 Then RsAnt.Update
      RsAtualizacao.MoveNext
      If a > TotalReg Then Exit Do
      a = a + 1
    Loop
    'Close #Fnum
Else
   Set RsAnt = AbreRecordset("select * from PRECOATUALI")

    If RsAnt.EOF Then
       MsgBox "Não Exite Arquivo Anterior Gravado...", 48, "Aviso"
       Exit Sub
    End If
    RsAnt.MoveLast
    LcTo = RsAnt.RecordCount
    RsAnt.MoveFirst
    a = 1
    Do Until RsAnt.EOF
       LcAchouEsp = True
       Me.Caption = "Atualizando Registro " & a & " de " & LcTo
       Set RsAtualizacao = AbreRecordset("Select * from Produtos where codigo=" & RsAnt!Codigo)
       If Not RsAtualizacao.EOF Then
         ' RsAtualizacao.Edit
          RsAtualizacao!Preco = RsAnt!Preco
          RsAtualizacao!LimiteVenda = RsAnt!PrecoLimiteVenda
          RsAtualizacao!MinimoVenda = RsAnt!MiminoVenda
          RsAtualizacao.Update
       End If
       RsAnt.MoveNext
       a = a + 1
    Loop
    
End If
RsAnt.MoveFirst
If RsAnt.EOF Then
   Arquivo.Enabled = False
Else
   Arquivo.Enabled = True
End If
RsAnt.Close
RsAtualizacao.Close
Set RsAtualizacao = Nothing
Set RsAnt = Nothing

Me.Caption = LcCapt

LcEscolha = 1
MsgBox "Operação Realizada com Sucesso !!!", 48, "Aviso"


End Sub

Private Sub CmdOk_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{F}"

End Sub

Private Sub Distribuidor_Click()
On Error GoTo Errd
LcEscolha = 5
AbreBanco (produtora)
RsAtual.Index = "Especie"
cbo.Clear
Do Until RsAtual.EOF
   cbo.AddItem RsAtual!Produtor
   RsAtual.MoveNext
Loop
RsAtual.Close
cbo.Enabled = True
Label1(0).Enabled = True
Label1(1).Enabled = True
txt.Enabled = True
option1.Enabled = True
Frame2.Enabled = True
cbo.SetFocus
Exit Sub
Errd:
Exit Sub
End Sub

Private Sub Form_Activate()
Set GlFormA = Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Teclas (KeyCode)
End Sub
Function Carregaforn()
On Error GoTo errc
Dim RsFornecedor As Recordset
AbreBase
LcSql = "Select * from ALID002 order by razaosoc"
Set RsFornecedor = Dbbase.OpenRecordset(LcSql, dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcTamanho = 0
fornecedor.Clear
Do Until RsFornecedor.EOF
   ReDim Preserve Mtfor(LcTamanho)
   Mtfor(LcTamanho).Codigo = RsFornecedor!Codigo
   Mtfor(LcTamanho).Nome = RsFornecedor!RazaoSoc
   fornecedor.AddItem RsFornecedor!RazaoSoc
   RsFornecedor.MoveNext
   LcTamanho = LcTamanho + 1
Loop
If LcTamanho > 0 Then LcTamanho = LcTamanho - 1
'Comissao.AddItem "TODOS"
'Comissao.Text = "TODOS"
RsFornecedor.Close
Set RsFornecedor = Nothing
Exit Function
errc:

Exit Function
End Function
Private Sub Form_Load()
On Error Resume Next
Dim RsAnt As ADODB.Recordset
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
Carregaforn
AbreBase
'LcNomeArquivos = App.Path & "\Reajuste.Rel"
'Set Dbbase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
Set RsAnt = AbreRecordset("select * from PRECOATUALI")

If RsAnt.EOF Then
   Arquivo.Enabled = False
Else
   Arquivo.Enabled = True
End If
LcEscolha = 2
LiberaDescricao
RsAnt.Close
Dbbase.Close

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
FrmPrincipal.SetFocus
End Sub

Private Sub Individual_Click()
'Label2.Visible = True
LiberaDescricao
End Sub

Sub LiberaDescricao()
LcEscolha = 2
Frame2.Enabled = True
bo.Enabled = True
'Label2.Visible = True
Reais.Enabled = True
Label1(0).Enabled = True
Label1(1).Enabled = True
Percentual.Enabled = True

txt.Enabled = True
bo.SetFocus
End Sub
Private Sub Individual_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{F}"

End Sub

Private Sub Option1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{F}"

End Sub

Private Sub Percentual_Click()
Label1(1).Caption = "Valor Percentual"
Label1(2).Visible = True
End Sub

Private Sub Percentual_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{F}"

End Sub

Private Sub Reais_Click()
Label1(1).Caption = "Novo Valor"
Label1(2).Visible = False
End Sub

Private Sub Reais_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{F}"

End Sub

Private Sub Todas_Click()
On Error Resume Next
LcEscolha = 1
bo.Enabled = False
Percentual.Value = True
Reais.Enabled = False
'Label2.Visible = False
Label1(0).Enabled = False
Label1(1).Enabled = True
txt.Enabled = True
option1.Enabled = True
Frame2.Enabled = True
bo.Text = ""
bo.Enabled = False
txt.SetFocus
End Sub
Function BuscaProduto()
On Error Resume Next
Dim LcDigitado As String
Dim LcAchou As Integer
Dim RsProduto As Recordset, RsUnidade As Recordset
If Len(bo.Text) = 0 Then Exit Function
AbreBase
LcDigitado = bo.Text
bo.Text = Right("00000" & bo.Text, 5)
Set RsProduto = Dbbase.OpenRecordset("select * From alid009 where cod='" & bo.Text & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic) ', dbOpenDynaset)
If Not RsProduto.EOF Then
    Codigo.Text = RsProduto!cod
    bo.Text = RsProduto!Nome
    LcAchou = True
Else
    bo.Text = LcDigitado
    GlCriterioSql = "select * From alid009 where nome like '" & UCase(bo.Text) & "*'  order by nome"
    FrmPesquisaProdutos.Show , Me
    LcAchou = True
  End If
If LcAchou Then txt.SetFocus
RsProduto.Close
Set RsProduto = Nothing

End Function

Private Sub Todas_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{F}"

End Sub

Private Sub Txt_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{F}"

End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
End Sub
