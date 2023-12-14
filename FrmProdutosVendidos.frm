VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FrmProdutosVendidos 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Produtos Vendidos"
   ClientHeight    =   3045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   ScaleHeight     =   3045
   ScaleWidth      =   6495
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox NCM 
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   480
      Width           =   2775
   End
   Begin VB.TextBox produto 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   5535
   End
   Begin VB.CommandButton CmdGerarRelatorio 
      Caption         =   "Gerar"
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   2295
   End
   Begin MSMask.MaskEdBox DataI 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "99/99/99"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox DataF 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "99/99/99"
      PromptChar      =   "_"
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "NCM"
      Height          =   255
      Left            =   3000
      TabIndex        =   8
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Msg 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   2400
      Width           =   6255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Produto"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "FrmProdutosVendidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Rel       As New Rel_Posicao_Vendas_Produto
Private Sub CmdGerarRelatorio_Click()
Dim LcCap As String
LcCap = Me.Caption
Me.MousePointer = 11
GeraPorNota
GeraPorVale
Imprime
Me.MousePointer = 0
Me.Caption = LcCap
End Sub
Sub Imprime()
Dim StrSql As String

Dim Rs As ADODB.Recordset


StrSql = "SELECT * from rel_vendas order by descricao"
'Debug.Print StrSql

Set Rs = AbreRecordset(StrSql)

Load Relatorios

With Relatorios
     Rel.DiscardSavedData
    
     Rel.Database.SetDataSource Rs
     .CRViewer1.ReportSource = Rel
     setaformula
      .CRViewer1.ViewReport
End With
Relatorios.Show

Screen.MousePointer = vbDefault

Me.Caption = LcCap

End Sub
Sub setaformula()
Dim a As Integer
Dim LcFormula, LcCriterio As String, LcTipoSaida As Integer
Dim RsEmpresa As Recordset
Dim RsOpcao As Recordset
Dim LcValor As Double
Dim LcEmpresa, LcEndereco, LcFone, LcCelular, Lccelular1, Lcemail, LcVer, LcCap, LcVer1 As String
Dim lctitulo As String
Dim StrSql As String
Dim bb     As Database

Set db = OpenDatabase(GLBase)
Set RsEmpresa = db.OpenRecordset("Select * from EMPRESA")

If Not RsEmpresa.EOF Then
   LcEmpresa = RsEmpresa!Razao & ""
   LcEndereco = RsEmpresa!Endereco & " " & RsEmpresa!Bairro
   LcFone = RsEmpresa!Fone & "" & IIf(Not IsNull(RsEmpresa!Fax), " Fax:" & RsEmpresa!Fax, "")
   LcInscricao = RsEmpresa!inscricaoestadual & ""
   LcCnpj = RsEmpresa!CGC & ""
   Celular = "Insc. Estadual: " & LcInscricao '.608783.0021'"
   Lcemail = "CNPJ: " & LcCnpj '.682.162/0001-88'"
End If
Set RsEmpresa = Nothing
lctitulo = "Produtos Vendidos no pediodo:" & DataI.Text & " a " & DataF.Text
If Len(produto) > 0 Then
   lctitulo = lctitulo & " Iniciados por:" & produto.Text
End If
With Rel
'Exit Sub
For a = 1 To .FormulaFields.Count
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("Fone") Then .FormulaFields(a).Text = "totext('" & LcFone & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("EMPRESA") Then .FormulaFields(a).Text = "totext('" & LcEmpresa & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("ENDERECO") Then .FormulaFields(a).Text = "totext('" & LcEndereco & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("EMAIL") Then .FormulaFields(a).Text = "totext('" & Lcemail & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("Celular") Then .FormulaFields(a).Text = "totext('" & Celular & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("msg1") Then .FormulaFields(a).Text = "totext('" & GlMsg & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("msg2") Then .FormulaFields(a).Text = "totext('" & GlMsg1 & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("msg3") Then .FormulaFields(a).Text = "totext('" & GlMsg2 & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("Titulo") Then
           .FormulaFields(a).Text = "totext('" & lctitulo & "')"
        End If
        
    Next
End With
End Sub
Sub GeraPorNota()
Dim Rs As New ADODB.Recordset
Dim RsUnidade As ADODB.Recordset
Dim StrSql As String
Dim TotalReg As Long
Dim RegistroAtual As Long
StrSql = "SELECT alid050.NUMNF,alid052.descricao,alid052.ClassificacaoFiscal,alid052.QTDE, Sum(QTDE*valunit) AS Valor_Unidade,Sum((alid052.QTDE*alid052.valunit)-alid052.Desconto) AS Valor_Total, alid052.codProd, produtos.QtdMedida, produtos.unidMedida, alid052.unimed, alid052.QTDUM "
StrSql = StrSql & " FROM alid050 INNER JOIN (alid052 INNER JOIN produtos ON alid052.codProd = produtos.codigo) ON alid050.NUMNF = alid052.NUMNF"
StrSql = StrSql & " WHERE (((alid050.DTEMIS) Between #" & Format(DataI.Text, "mm/dd/yy") & "# And #" & Format(DataF.Text, "mm/dd/yy") & "#) and ((alid050.NATUREZA)<>'TRANSFERENCIA' And (alid050.NATUREZA)<>'ENTRADA') AND ((alid050.status)='Autorizado o uso da NF-e')"
If Len(NCM.Text) > 0 Then
   StrSql = StrSql & " and alid052.ClassificacaoFiscal like'%" & Replace(NCM.Text, "'", "''") & "%')"
Else
   StrSql = StrSql & ")"
End If

If Len(produto.Text) > 0 Then StrSql = StrSql & " and descricao like '" & produto.Text & "%'"
StrSql = StrSql & "  GROUP BY alid050.NUMNF,alid052.descricao,alid052.ClassificacaoFiscal,alid052.unimed, alid052.QTDE,alid052.valunit, alid052.QTDUM, alid052.codProd, produtos.QtdMedida,produtos.unidMedida;"
Debug.Print StrSql
Set Rs = AbreRecordset(StrSql, True)
'==> Apaga as informaçoes antiga
ExecutaSql "Delete from rel_vendas"
'===> Recupera total de Registros
If Not Rs.EOF Then
   Rs.MoveLast
   TotalReg = Rs.RecordCount
   Rs.MoveFirst
End If
Msg.Caption = ""
Do Until Rs.EOF
   DoEvents
   RegistroAtual = RegistroAtual + 1
   Msg.Caption = "Registro " & RegistroAtual & " de " & TotalReg & " Nº NF: " & Rs!NumNf & " Prod:" & Rs!codProd & " - " & Rs!Descricao
   DoEvents
   Dim LcUnidadeProduto As String
   Dim LcUnidadeVenda As String
   Set RsUnidade = AbreRecordsetRel("Select * from Alid004 where simbolo='" & Rs!UNIMED & "'", RsUnidade)
   If Not RsUnidade.EOF Then
         LcUnidadeProduto = RsUnidade!Simbolo & " c/ " & Rs!QtdMedida
   End If
   Dim QtdCX As Currency
   Dim Resto As Currency
   Dim QtdMed As Currency
   QtdMed = 1
   If (Rs!UnidMedida = RsUnidade!cod) And (Rs!QtdMedida = Rs!QTDUM) Then
      '===>Vendeu o padrao
      QtdCX = Rs!Qtde
      Resto = 0
   Else
      QtdCX = 0
      Resto = Rs!Qtde
   End If
   
   'If IsNumeric(Rs!QtdMedida) Then
    '  If Rs!QtdMedida > 0 Then QtdMed = Rs!QtdMedida
   'End If
   'QtdCX = Int(Rs!Valor_Unidade / QtdMed)
   'Resto = Rs!Valor_Unidade Mod QtdMed
   LcUnidadeVenda = Rs!UNIMED & " C/ " & Rs!QTDUM
   '==> Inclui os registros.
   StrSql = "Insert into rel_vendas (CodProduto,Descricao,Quantidade_Unidade,Quantidade_CX,Cx_Unidade,Unidade,UnidadeVenda,NCM,valor) values("
   StrSql = StrSql & Rs!codProd & ","
   StrSql = StrSql & "'" & Rs!Descricao & "',"
   StrSql = StrSql & Replace(Replace(Rs!Valor_Unidade, ".", ""), ",", ".") & ","
   StrSql = StrSql & Replace(Replace(QtdCX, ".", ""), ",", ".") & ","
   StrSql = StrSql & Replace(Replace(Resto, ".", ""), ",", ".") & ","
   StrSql = StrSql & "'" & LcUnidadeProduto & "',"
   StrSql = StrSql & "'" & LcUnidadeVenda & "',"
   StrSql = StrSql & "'" & Rs!ClassificacaoFiscal & "',"
   StrSql = StrSql & Replace(Replace(Rs!Valor_Total, ".", ""), ",", ".") & ")"
   ExecutaSql StrSql
   'Debug.Print StrSql
   'Debug.Print DEscricaoErro
   Rs.MoveNext
Loop

End Sub
Sub GeraPorVale()
Dim Rs As New ADODB.Recordset
Dim RsUnidade As ADODB.Recordset
Dim StrSql As String
Dim TotalReg As Long
Dim RegistroAtual As Long

StrSql = "SELECT vales.NUMNF,valesprodutos.codProd, valesprodutos.descricao, Sum(QTDE*QTDUM) AS Valor_Unidade,Sum((QTDE*valunit)-Desconto) AS Valor_Total,produtos.classificacaofiscal, produtos.QtdMedida, produtos.unidMedida, valesprodutos.unimed, valesprodutos.QTDUM"
StrSql = StrSql & " FROM (vales INNER JOIN valesprodutos ON vales.NUMNF = valesprodutos.NUMNF) INNER JOIN produtos ON valesprodutos.codProd = produtos.codigo"
StrSql = StrSql & " WHERE (((vales.DTEMIS) Between #" & Format(DataI.Text, "mm/dd/yy") & "# And #" & Format(DataF.Text, "mm/dd/yy") & "#) AND ((vales.baixado)=0)"
If Len(NCM.Text) > 0 Then
   StrSql = StrSql & " and produtos.classificacaofiscal like '%" & Replace(NCM.Text, "'", "''") & "%')"
Else
   StrSql = StrSql & ")"
End If
If Len(produto.Text) > 0 Then StrSql = StrSql & " and descricao like '" & produto.Text & "%'"
StrSql = StrSql & " GROUP BY vales.NUMNF,valesprodutos.codProd,valesprodutos.valunit, valesprodutos.descricao, valesprodutos.QTDE, valesprodutos.QTDUM,produtos.classificacaofiscal, produtos.QtdMedida,produtos.unidMedida;"

Set Rs = AbreRecordset(StrSql, True)
'==> Apaga as informaçoes antiga
If Not Rs.EOF Then
   Rs.MoveLast
   TotalReg = Rs.RecordCount
   Rs.MoveFirst
End If
Msg.Caption = ""

Do Until Rs.EOF
   Dim QtdCX As Currency
   Dim Resto As Currency
   Dim LcUnidadeProduto As String
   Dim LcUnidadeVenda As String
    DoEvents
   RegistroAtual = RegistroAtual + 1
   Msg.Caption = "Registro " & RegistroAtual & " de " & TotalReg & " Nº Vale: " & Rs!NumNf & " Prod:" & Rs!codProd & " - " & Rs!Descricao
   DoEvents
   
   Set RsUnidade = AbreRecordsetRel("Select * from Alid004 where cod='" & Rs!UnidMedida & "'", RsUnidade)
   If Not RsUnidade.EOF Then
         LcUnidadeProduto = RsUnidade!Simbolo & " c/ " & Rs!QtdMedida
   End If
   Dim QtdMed As Currency
   QtdMed = 1
   If IsNumeric(Rs!QtdMedida) Then
      If Rs!QtdMedida > 0 Then QtdMed = Rs!QtdMedida
   End If
   QtdCX = Int(Rs!Valor_Unidade / QtdMed)
   Resto = Rs!Valor_Unidade Mod QtdMed
   LcUnidadeVenda = Rs!UNIMED & " C/ " & Rs!QTDUM
   '==> Inclui os registros.
   StrSql = "Insert into rel_vendas (CodProduto,Descricao,Quantidade_Unidade,Quantidade_CX,Cx_Unidade,Unidade,UnidadeVenda,NCM,valor) values("
   StrSql = StrSql & Rs!codProd & ","
   StrSql = StrSql & "'" & Rs!Descricao & "',"
   StrSql = StrSql & Replace(Replace(Rs!Valor_Unidade, ".", ""), ",", ".") & ","
   StrSql = StrSql & Replace(Replace(QtdCX, ".", ""), ",", ".") & ","
   StrSql = StrSql & Replace(Replace(Resto, ".", ""), ",", ".") & ","
    StrSql = StrSql & "'" & LcUnidadeProduto & "',"
   StrSql = StrSql & "'" & LcUnidadeVendao & "',"
   StrSql = StrSql & "'" & Rs!ClassificacaoFiscal & "',"
    StrSql = StrSql & Replace(Replace(Rs!Valor_Total, ".", ""), ",", ".") & ")"
   ExecutaSql StrSql
   Rs.MoveNext
Loop
End Sub
Function AbreRecordsetRel(LcSql As String, RsAtual As ADODB.Recordset) As ADODB.Recordset

On Error GoTo ErroAbreRs
LcComentario = "- AbreRecordset - Criando Nova Instancia do RecordSet."
Dim conexao As New ADODB.Connection
Dim strConnect As String
Set RsAtual = New ADODB.Recordset
LcComentario = "- AbreRecordset - Setando os Parametros do Recordset."
RsAtual.CursorType = adOpenDynamic ' adOpenStatic
RsAtual.CursorLocation = adUseClient
RsAtual.LockType = adLockReadOnly
RsAtual.Source = LcSql
strConnect = "driver={Microsoft Access Driver (*.mdb)};DBQ=" & GLBase & ";UID=Admin;PWD=;"
Set conexaoo = New ADODB.Connection
conexao.CursorLocation = adUseClient
'usamos um cursor do lado do cliente pois os dados 'serao acessados na maquina do cliente e nao de um servidor
LcComentario = "- Função 'abreconexao - Abrindo a Conexão com o DB."
'MsgBox strConnect
conexao.Open strConnect

RsAtual.ActiveConnection = conexao

LcComentario = "- AbreRecordset - Abrindo o Recordset."
RsAtual.Open
Set AbreRecordsetRel = RsAtual
Exit Function

ErroAbreRs:
'If err.Number = 3709 Then
'   'abreconexao
'   Resume 0
'End If
'If LcExibemsg Then ErrosSistema = MsgBox(msg, 64, "erro Abrindo Tabela. ") Else ErrosSistema = 0
'MsgBox err.Description & err.Number
'Resume 0
logErro err.Number, err.Description, LcComentario
Resume Next
End Function
