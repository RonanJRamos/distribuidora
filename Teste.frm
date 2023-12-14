VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form Teste 
   Caption         =   "Form1"
   ClientHeight    =   5520
   ClientLeft      =   135
   ClientTop       =   510
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   ScaleHeight     =   5520
   ScaleWidth      =   6105
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdEvitaDuplicado 
      Caption         =   "Evita Duplicado Pedido"
      Height          =   615
      Left            =   0
      TabIndex        =   14
      Top             =   4080
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Verifica NFe Lancada Contas a Receber"
      Height          =   615
      Left            =   2880
      TabIndex        =   13
      Top             =   4560
      Width           =   2775
   End
   Begin VB.CommandButton CmdAcertaCreditoCliente 
      Caption         =   "Acerta Credito Utilizado do Cliente"
      Height          =   495
      Left            =   2880
      TabIndex        =   12
      Top             =   2760
      Width           =   2775
   End
   Begin VB.CommandButton CmdLancaCodigoNota 
      Caption         =   "Lanca Codigo NF Entrada"
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   2760
      Width           =   2535
   End
   Begin VB.CommandButton CmdAcertaSaldo 
      Caption         =   "AcertaSaldo"
      Height          =   375
      Left            =   2880
      TabIndex        =   10
      Top             =   4080
      Width           =   2775
   End
   Begin MSMask.MaskEdBox Dataf 
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   3720
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Datai 
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin VB.CommandButton CmdAcertaIpi 
      Caption         =   "Lança Estoque Fiscal"
      Height          =   495
      Left            =   2880
      TabIndex        =   6
      Top             =   3480
      Width           =   2775
   End
   Begin VB.CommandButton AcertaImpostoSaida 
      Caption         =   "Acerta o Imposto de Saida"
      Height          =   615
      Left            =   2880
      TabIndex        =   5
      Top             =   2040
      Width           =   2775
   End
   Begin VB.CommandButton CmdAcertaData 
      Caption         =   "&Acertar data de NF Saida."
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   2535
   End
   Begin VB.CommandButton CmdGerarDadosSintegra 
      Caption         =   "Gerar Dados Sintegra"
      Height          =   735
      Left            =   2880
      TabIndex        =   3
      Top             =   1200
      Width           =   2775
   End
   Begin VB.CommandButton CmdVerificarSintegra 
      Caption         =   "Verifcar dados Sintegra"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Acertar os Itens da nota Fiscal de entrada com fornecedor e data"
      Height          =   855
      Left            =   2880
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Lança os custos da nota de entrada para os produtos."
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Periodo"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6120
      Y1              =   3360
      Y2              =   3360
   End
End
Attribute VB_Name = "Teste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ValorBase   As Double
Private valorIcms   As Double
Private Rs52        As ADODB.Recordset
Private RsProduto   As ADODB.Recordset



Private Sub AcertaImpostoSaida_Click()
'On Error Resume Next
Dim Rs50     As ADODB.Recordset
Dim StrSql As String
Dim a As Long
Dim TotalReg As Long
Dim Codigo As Long
StrSql = "Select * from alid050 where dtemis>='2004-01-01' and valoricms=0 order by numnf"

Set Rs50 = AbreRecordset(StrSql, True)
StrSql = "Select * from alid052 order by numnf"
Set Rs52 = AbreRecordset(StrSql, True)
StrSql = "Select * from produtos order by codigo"
Set RsProduto = AbreRecordset(StrSql, True)

If Not Rs50.EOF Then
   Rs50.MoveLast
   TotalReg = Rs50.RecordCount
   Rs50.MoveFirst
End If
a = 0
Do Until Rs50.EOF
  a = a + 1
  Me.Caption = "Processando nota " & Rs50!NUMNF & " - Registro:" & a & " de " & TotalReg
  DoEvents
  BuscaValor Rs50!NUMNF
  Codigo = Rs50!Codigo
  'Rs50!BaseIcms = AcertaNumero(CStr(ValorBase), 2)
  'Rs50!ValorIcms = AcertaNumero(CStr(ValorIcms), 2)
  'Rs50.Update
  StrSql = "Update alid050 set baseicms=" & Replace(ValorBase, ",", ".") & ","
  StrSql = StrSql & " ValorIcms=" & Replace(valorIcms, ",", ".")
  StrSql = StrSql & " where codigo=" & Codigo
  
  afetados = ExecutaSql(StrSql)
  Rs50.MoveNext
Loop

MsgBox "Processo terminado..", 64, "Aviso"

End Sub
Sub BuscaValor(Nota As String)
'On Error Resume Next
Dim StrSql As String
Dim ValorProduto As Double
Dim icms As Double
ValorBase = 0
valorIcms = 0
Rs52.MoveFirst

Rs52.Find "numnf='" & Nota & "'"
RsProduto.MoveFirst

Do Until Rs52.EOF
   '==>Verifica o icms
   If Rs52!NUMNF <> Nota Then Exit Do
   If Not IsNull(Rs52!icms) Then
        If Rs52!icms > 0 Then
           ValorProduto = Rs52!VALUNIT * Rs52!QTDE
           ValorBase = ValorBase + ValorProduto
           valorIcms = valorIcms + (ValorProduto * (Rs52!icms / 100))
        End If
   Else
   RsProduto.Find "Codigo=" & Rs52!codProd
     If Not RsProduto.EOF Then
        'If RsProduto!Icms = 0 Or IsNull(RsProduto!Icms) Then
           If Val(RsProduto!cst) = 60 Or Val(RsProduto!cst) = 160 Or Val(RsProduto!cst) = 260 Then
              icms = 0
           Else
              icms = 18
           End If
        'Else
        '   Icms = RsProduto!Icms
        ' End If
     End If
         If icms > 0 Then
           ValorProduto = Rs52!VALUNIT * Rs52!QTDE
           ValorBase = ValorBase + ValorProduto
           valorIcms = valorIcms + (ValorProduto * (icms / 100))
         End If
   End If
   Rs52.MoveNext
Loop
'Set Rs52 = Nothing
End Sub

Private Sub CmdAcertaCreditoCliente_Click()
Dim RsConta As ADODB.Recordset
Dim db As Database
Dim StrSql As String
Dim TotalReg As Long
Dim LcCap As String
Dim a As Long

StrSql = "SELECT alid015.CLIENTE, Sum(alid015.VALOR) AS SomaDeVALOR From alid015 "
StrSql = StrSql & "Where (((alid015.VALPAGO) = '0')) "
StrSql = StrSql & "GROUP BY alid015.CLIENTE "
StrSql = StrSql & "Having (((alid015.Cliente) <> '')) "
StrSql = StrSql & "ORDER BY alid015.CLIENTE;"
LcCap = Me.Caption
Screen.MousePointer = 11

Set RsConta = AbreRecordset(StrSql, True)
RsConta.MoveLast
TotalReg = RsConta.RecordCount
RsConta.MoveFirst
Set db = OpenDatabase(GLBase)
Do Until RsConta.EOF
   a = a + 1
   Me.Caption = "Processando registro " & a & " de " & TotalReg
   DoEvents
   StrSql = "Update alid001 set CreditoUtilizado=" & Replace(RsConta!SomaDeVALOR, ",", ".") & " Where CODIGO='" & Right("00000" & RsConta!Cliente, 5) & "'"
   db.Execute StrSql
   RsConta.MoveNext
Loop
Screen.MousePointer = 0
MsgBox "Opereção Terminada.", 64, "Aviso"



End Sub

Private Sub CmdAcertaData_Click()
Dim Rs50    As ADODB.Recordset
Dim Rs54    As ADODB.Recordset
Dim RsNota  As ADODB.Recordset
Dim RsGeral As ADODB.Recordset

Dim StrSql  As String
On Error GoTo Errogeral
'==Abre a nota fiscal
LcCap = Me.Caption
Me.Caption = "Abrindo o banco de notas..."
DoEvents
conexaoAdo.BeginTrans

StrSql = "Select * from alid050 order by numnf;"
Set RsNota = AbreRecordset(StrSql, True)


'===> Acerta no reg 50
StrSql = "select * from sintegra_50 where (cfop like '5%') or (cfop like '6%')order by nf;"
Set Rs50 = AbreRecordset(StrSql, True)

Do Until Rs50.EOF
  Me.Caption = "Acertando Registro 50 da nota " & Rs50!NF
  DoEvents
  '==> Pesquisa a nota no cadastro de notas
 ' If Rs50!NF = "055502" Then Stop
  RsNota.MoveFirst
  RsNota.Find "numnf='" & Rs50!NF & "'"
  If Not RsNota.EOF Then
     afetados = ExecutaSql("Update sintegra_50 Set data='" & Format(RsNota!DTEMIS, "yyyy-mm-dd") & "' where nf='" & Rs50!NF & "'")
    ' Rs50.Update
  End If
  Rs50.MoveNext
Loop
Set Rs50 = Nothing


'===> Acerta no reg 54
StrSql = "select * from sintegra_54 where (cfop like '5%') or (cfop like '6%')order by nf;"
Set Rs54 = AbreRecordset(StrSql, True)
If Not Rs54.EOF Then
   Rs54.MoveLast
   Total_Reg = Rs54.RecordCount
   Rs54.MoveFirst
End If

Do Until Rs54.EOF
  a = a + 1
  Me.Caption = "Acertando Registro 54 da nota " & Rs54!NF & " Reg." & a & " de " & Total_Reg
  DoEvents
  '==> Pesquisa a nota no cadastro de notas
  RsNota.MoveFirst
  RsNota.Find "numnf='" & Rs54!NF & "'"
  If Not RsNota.EOF Then
     afetados = ExecutaSql("Update sintegra_54 Set data='" & Format(RsNota!DTEMIS, "yyyy-mm-dd") & "' where nf='" & Rs54!NF & "'")
  End If
  Rs54.MoveNext
Loop
Set Rs54 = Nothing


'===> Acerta no reg Sintegra
StrSql = "select * from sintegra where (cfop like '5%') or (cfop like '6%')order by nf;"
Set RsGeral = AbreRecordset(StrSql, True)
If Not RsGeral.EOF Then
   RsGeral.MoveLast
   Total_Reg = RsGeral.RecordCount
   RsGeral.MoveFirst
End If
Do Until RsGeral.EOF
  a = a + 1
  Me.Caption = "Acertando Registro do Sintegra da nota " & RsGeral!NF & " Reg " & a & " de " & Total_Reg
  DoEvents
  '==> Pesquisa a nota no cadastro de notas
  RsNota.MoveFirst
  RsNota.Find "numnf='" & RsGeral!NF & "'"
  If Not RsNota.EOF Then
     afetados = ExecutaSql("Update sintegra Set data='" & Format(RsNota!DTEMIS, "yyyy-mm-dd") & "' where nf='" & Sintegra!NF & "'")
  End If
  RsGeral.MoveNext
Loop
Set RsGeral = Nothing
conexaoAdo.CommitTrans
Me.Caption = LcCap
MsgBox "Terminei."
Exit Sub
Errogeral:

conexaoAdo.RollbackTrans
MsgBox "Ocorreu erro " & err.Number & " " & err.Description
Resume 0
End Sub

Private Sub CmdAcertaIpi_Click()
'==> Esta função busca as notas de entrada e saida desde 01-01-04
'==> e lança no estoque fiscal, as quantidades tem que ser convertidas.
'==> para a unidade principal
Dim RsSaida     As ADODB.Recordset
Dim RsEntrada   As ADODB.Recordset
Dim RsProduto   As ADODB.Recordset
Dim RsUnid      As Recordset
Dim StrSql      As String
Dim a           As Long
Dim TotalRegs   As Long
Dim QUnitario   As Double
Dim QBase       As Double
Dim CustoUnit   As Double
Dim NovocustoUnit As Double
Dim db          As Database

LcCap = Me.Caption
Me.Caption = "Abrindo a tb notas de entrada."
DoEvents
'==> Abre as entradas

StrSql = "SELECT itensentradanf.*,entradanf.processado,entradanf.codigo as Codigoentrada, entradanf.DATA as entrada " & _
         "FROM entradanf INNER JOIN itensentradanf ON entradanf.NF = itensentradanf.NUMNF " & _
         "Where entradanf.DATA Between '" & Format(Datai.Text, "yyyy-mm-dd") & "' And '" & Format(Dataf.Text, "yyyy-mm-dd") & "' and entradanf.Processado=0 " & _
         "order by entradanf.codigo;"
StrSql = "select * from itensentradanf order by codigo"

Set RsEntrada = AbreRecordset(StrSql, True)

Me.Caption = "Abrindo a tb notas de Saida."
DoEvents

StrSql = "SELECT alid052.*, alid050.Processado,alid050.codigo as CodigoSaida, alid050.DTEMIS FROM alid050 INNER JOIN alid052 ON alid050.NUMNF = alid052.NUMNF " & _
         "WHERE alid050.Dtemis Between '" & Format(Datai.Text, "yyyy-mm-dd") & "' And '" & Format(Dataf.Text, "yyyy-mm-dd") & "' and alid050.Processado=0 order by codigo;"

'Set RsSaida = AbreRecordset(StrSql, True)

Me.Caption = "Abrindo a tb Produtos."
DoEvents
Set RsProduto = AbreRecordset("Select * from produtos order by codigo", True)


'==> Processando os dados de entrada
If Not RsEntrada.EOF Then
  RsEntrada.MoveLast
  TotalRegs = RsEntrada.RecordCount
  RsEntrada.MoveFirst
End If

a = 0

Do Until RsEntrada.EOF
   a = a + 1
   Me.Caption = "Processando Entrada reg. " & a & " de " & TotalRegs
   DoEvents
   '==>Compara as unidades
   RsProduto.Filter = "Codigo=" & RsEntrada!Item
   Codigoentrada = RsEntrada!Codigo
   If Not RsProduto.EOF Then
      'If RsEntrada!Unimed <> RsProduto!UnidMedida And RsEntrada!QTDUM <> RsProduto!QtdMedida Then
         '==> Recupera na menor unidade
         QBase = RsEntrada!QTDE * RsEntrada!QTDUM
         QBase = QBase + RsProduto!QuantEstoque
         '==> tranforma para a unidade de cadastro
         'QBase = QUnitario / RsProduto!QtdMedida
         
         '==> Acharemos o novo valor
         
         CustoUnit = RsEntrada!VALUNIT
         'NovocustoTotal = NovocustoUnitario * QBase
      'Else
      '   QBase = RsEntrada!Qtde
      '   NovocustoUnit = RsEntrada!VALUNIT
      'End If
   Else
       QBase = RsEntrada!QTDE * RsEntrada!QTDUM
       NovocustoUnit = RsEntrada!VALUNIT
   End If
  ' EntradaFiscal RsEntrada!item, QBase, RsEntrada!Entrada, RsEntrada!Descricao, NovocustoUnit, RsEntrada!ValorTotal
   StrSql = "UPdate produtos set QuantEstoque=" & Replace(CStr(QBase), ",", ".") & ",Santa1=" & Replace(CStr(QBase), ",", ".") & " Where codigo=" & RsEntrada!Item
   
   'StrSql = "Update entradanf Set Processado=1 where Codigo=" & RsEntrada!Codigoentrada
   afetados = ExecutaSql(StrSql)
   RsEntrada.MoveNext
Loop

Set RsEntrada = Nothing
MsgBox "Estoque fiscal gerado com sucesso."
Exit Sub
a = 0
If Not RsSaida.EOF Then
  RsSaida.MoveLast
  TotalRegs = RsSaida.RecordCount
  RsSaida.MoveFirst
End If
Set db = OpenDatabase(GLBase)

Do Until RsSaida.EOF
  a = a + 1
   Me.Caption = "Processando Saida reg. " & a & " de " & TotalRegs
   DoEvents
   '==>Compara as unidades
   'Set RsProduto = AbreRecordsetLeitura("Select * from produtos order by codigo")

   RsProduto.Filter = "Codigo=" & RsSaida!codProd
   
   If Not RsProduto.EOF Then
      '==> Abre a tb de unidade
      Set RsUnidade = db.OpenRecordset("Select * from alid004 where cod='" & RsProduto!unidMedida & "'")
      
      If RsSaida!UNIMED <> RsUnidade!Simbolo And RsSaida!QTDUM <> RsProduto!QtdMedida Then
         '==> Recupera na menor unidade
         QUnitario = RsSaida!QTDE * IIf(RsSaida!QTDUM > 0, RsSaida!QTDUM, 1)
         '==> tranforma para a unidade de cadastro
         QBase = QUnitario / RsProduto!QtdMedida
         
         '==> Acharemos o novo valor
         CustoUnit = (RsSaida!VALUNIT * RsSaida!QTDE) / QUnitario
         NovocustoUnit = CustoUnit * RsProduto!QtdMedida
         'NovocustoTotal = NovocustoUnitario * QBase
      Else
         QBase = RsSaida!QTDE
         NovocustoUnit = RsSaida!VALUNIT
      End If
   Else
       QBase = RsSaida!QTDE
       NovocustoUnit = RsSaida!VALUNIT
   End If
   CodigoSaida = RsSaida!Codigo
   SaidaFiscal RsSaida!codProd, QBase, RsSaida!DTEMIS, RsSaida!Descricao
   StrSql = "Update Alid050 Set Processado=1 where Codigo=" & RsSaida!CodigoSaida
   afetados = ExecutaSql(StrSql)

  RsSaida.MoveNext
Loop
Set RsSaida = Nothing
Call CmdAcertaSaldo_Click
MsgBox "Estoque fiscal gerado com sucesso."
End Sub

Private Sub CmdAcertaSaldo_Click()
'On Error Resume Next
Dim Rs As ADODB.Recordset
Dim Saldo As Double
Dim StrSql As String
Dim Codigo As Long
Dim TotalReg As Long
Dim a As Long
a = 0
StrSql = "SELECT * From estoquefiscal ORDER BY  CodigoProduto,data,codigo;"

Set Rs = AbreRecordset(StrSql, True)
Codigo = 0
If Not Rs.EOF Then
   Rs.MoveLast
   TotalReg = Rs.RecordCount
   Rs.MoveFirst
End If
Do Until Rs.EOF
   a = a + 1
   Me.Caption = "Reg " & a & " de " & TotalReg
   DoEvents
   If Codigo <> CLng(Rs!codigoproduto) Then
      Saldo = 0
      Codigo = CLng(Rs!codigoproduto)
   End If
   Saldo = Saldo + IIf(Not IsNull(Rs!Quantidade), Rs!Quantidade, 0) - IIf(Not IsNull(Rs!quantidadeSaida), Rs!quantidadeSaida, 0)
   'Rs!Saldo = Saldo
   'Rs.Update
   StrSql = "update estoquefiscal Set saldo=" & Replace(Saldo, ",", ".") & " where codigo=" & Rs!Codigo
   ExecutaSql StrSql
   Rs.MoveNext
Loop
MsgBox "Fim"

End Sub

Private Sub CmdEvitaDuplicado_Click()

Dim Rs As Recordset

Dim StrSql As String
Dim LcNFe As String
Dim LcCOdProd As Long
Dim LcUn As String
Dim a As Long
Dim Resposta As Integer
StrSql = "SELECT DISTINCTROW subproposta.NUMNF, subproposta.codProd, subproposta.ITEM, subproposta.QTDE, subproposta.VALUNIT, subproposta.UNIMED, subproposta.codigo, subproposta.descricao"
StrSql = StrSql & " From subproposta WHERE (((subproposta.NUMNF) In (SELECT [NUMNF] FROM [subproposta] As Tmp GROUP BY [NUMNF],[codProd] HAVING Count(*)>1  And [codProd] = [subproposta].[codProd])))"
StrSql = StrSql & " ORDER BY subproposta.NUMNF, subproposta.codProd,subproposta.UNIMED,subproposta.codigo;"

Set Dbbase = OpenDatabase(GLBase)
Debug.Print GLBase

Set Rs = Dbbase.OpenRecordset(StrSql)
Debug.Print StrSql
Debug.Print DEscricaoErro
Screen.MousePointer = 11
LcCap = Me.Caption
If Not Rs.EOF Then
   Rs.MoveLast
   total = Rs.RecordCount
   Rs.MoveFirst
End If
a = 1
Do Until Rs.EOF
   Me.Caption = "Efetuando correção Registro " & a & " de " & total
   DoEvents
   If Rs!NUMNF = LcNFe And Rs!codProd = LcCOdProd And Rs!UNIMED = LcUn Then
       StrSql = "Delete from subproposta where codigo=" & Rs!Codigo
       
        Rs.Delete
        
   Else
       LcNFe = Rs!NUMNF
       LcCOdProd = Rs!codProd
       LcUn = Rs!UNIMED
   End If
   
   Rs.MoveNext
   a = a + 1
Loop
'==> Criar o indice
Rs.Close
StrSql = "CREATE INDEX idx_Chave ON Subproposta (NUMNF, codProd,UNIMED) WITH PRIMARY"
 Dbbase.Execute StrSql
Screen.MousePointer = 0
MsgBox "Terminei"
End Sub

Private Sub CmdGerarDadosSintegra_Click()
Sintegra_gera_dados.Show , Me
End Sub

Private Sub CmdLancaCodigoNota_Click()
Dim Rs As ADODB.Recordset
Dim StrSql As String

abreconexao

Set Rs = AbreRecordset("Select * from entradanf order by nf,data", True)
Do Until Rs.EOF
  Me.Caption = "Processando " & Rs!NF
  DoEvents
  StrSql = "Update ItensEntradaNf Set " & _
           "codigonota=" & Rs!Codigo & _
           " where numnf='" & Rs!NF & "' and data='" & Format(Rs!Data, "yyyy-mm-dd") & "'"
           
  afetados = ExecutaSql(StrSql)
  afetados = ExecutaSql("Update Entradanf set nf='" & Right("00000" & Rs!NF, 6) & "' where codigo=" & Rs!Codigo)
  afetados = ExecutaSql("Update ItensEntradaNf Set NUMNF='" & Right("00000" & Rs!NF, 6) & "' where codigonota=" & Rs!Codigo)
  'If Afetados > 0 Then Stop
  Rs.MoveNext
Loop


StrSql = "Update entradanf Set " & _
        "Frete=0," & _
        "seguro=0," & _
        "PIS_COFINS=0," & _
        "NaoTributado=0," & _
        "DespesasAcessorias=0," & _
        "TipoFrete=2," & _
        "emissao=data"
ExecutaSql StrSql
MsgBox "Terminei"


End Sub

Private Sub CmdVerificarSintegra_Click()
Sintegra_Verifica.Show , Me
End Sub

Private Sub Command1_Click()
Dim Rs As ADODB.Recordset
Dim LcSql As String
Dim LcValor As String
Dim a As Long
LcCap = Me.Caption
Screen.MousePointer = 11
LcSql = "Select * from itensentradanf order by codigo"
Set Rs = AbreRecordset(LcSql)
Rs.MoveLast
LcTotal = Rs.RecordCount
Rs.MoveFirst
a = 1
Do Until Rs.EOF
   Me.Caption = "Efetuando atualização " & a & " de " & LcTotal
   DoEvents
   If Not IsNull(Rs!Item) Then
    If IsNumeric(Rs!VALUNIT) Then LcValor = Rs!VALUNIT Else LcValor = 0
    LcValor = Replace(LcValor, ",", ".")
    LcSql = "Update produtos SET custo=" & LcValor & " where codigo=" & Rs!Item
    lctotaql = ExecutaSql(LcSql)
    'If lctotaql > 0 Then Stop
    
   End If
   a = a + 1
   Rs.MoveNext
Loop
Screen.MousePointer = 0
MsgBox "Terminei"


End Sub

Private Sub Command2_Click()
On Error GoTo Errp
Dim Rs As ADODB.Recordset
Dim Rsp As Recordset
Dim a As Integer
Set Rs = AbreRecordset("Select * from entradanf order by nf", True)
Screen.MousePointer = 11
LcCap = Me.Caption
If Not Rs.EOF Then
   Rs.MoveLast
   total = Rs.RecordCount
   Rs.MoveFirst
End If
AbreBase
Set Rsp = Dbbase.OpenRecordset("Select * from alid002")
Do Until Rs.EOF
  If err.Number > 0 Then Exit Do
  Me.Caption = "PRocessando registro:" & a & " de " & total & " Nf=" & Rs!NF
  DoEvents
  Rsp.FindFirst "codigo='" & Rs!clicred & "'"
  If Not Rsp.EOF Then
     LcFor = Rsp!RazaoSoc & ""
  Else
     LcFor = ""
  End If
  LcSql = "Update itensentradanf SET fornecedor='" & LcFor & "', data='" & Format(Rs!Data, "yyyy-mm-dd") & "' where numnf='" & Rs!NF & "'"
  resp = ExecutaSql(LcSql)
  a = a + 1
  Rs.MoveNext
Loop
Screen.MousePointer = 0
Me.Caption = LcCap
MsgBox "Terminei"
Exit Sub
Errp:
MsgBox err.Description & err.Number
Resume Next
End Sub




Private Sub Command3_Click()
On Error GoTo Errp
Dim Rs As ADODB.Recordset
Dim Rsp As ADODB.Recordset
Dim a As Integer
Set Rs = AbreRecordset("Select * from alid050 where DTEMIS > '2011-01-01' order by numnf", True)
'MsgBox DEscricaoErro
Screen.MousePointer = 11
LcCap = Me.Caption
If Not Rs.EOF Then
   Rs.MoveLast
   total = Rs.RecordCount
   Rs.MoveFirst
End If
ExxibeContasAcertadas.Show , Me
Do Until Rs.EOF
  'If err.Number > 0 Then
   '  MsgBox err.Description & " - " & Rs!numnf
   '  Exit Do
  'End If
  lcnum = Rs("NumNf")
  Set Rsp = AbreRecordset("Select * from alid015 where nf like '" & lcnum & "%' and data>'2011-01-01'", True)
  Me.Caption = "Verificando NFE:" & a & " de " & total & " Nf=" & Rs!NUMNF
  DoEvents
  
  If Rsp.EOF Then
    '==> Vamos Lançar
    LancaConta Rs
   ' If err.Number > 0 Then
   '    MsgBox err.Description & " - " & Rs!numnf
   '    Exit Do
   ' End If
    'ExxibeContasAcertadas.List1.AddItem ("Lancada conta p/:" & Rs!numnf & " Emitida em: " & Rs!DTEMIS)
    DoEvents
  End If
  'If err.Number > 0 Then
  '   MsgBox err.Description & " - " & Rs!numnf
  '   Exit Do
  'End If
  a = a + 1
  Rs.MoveNext
Loop
MsgBox "Terminado"
Screen.MousePointer = 0
Me.Caption = LcCap
'MsgBox "Terminei"
Exit Sub
Errp:
MsgBox err.Description & err.Number
Resume Next
End Sub
Sub LancaConta(RsNota As ADODB.Recordset)
Dim RsTipoMonetario As Recordset

Dim RsContasReceber As ADODB.Recordset

LcSql3 = "Select * from Alid008"
AbreBase
Set RsTipoMonetario = Dbbase.OpenRecordset(LcSql3, dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcSql = "select * from Alid015 where NF like '" & RsNota!NUMNF & "%'"
Set RsContasReceber = AbreRecordset(LcSql)
LcNovo = True
If RsNota!condpag = "0 - A Vista" Then

           LcCriterioPes = "XTPMONET='" & RsNota!FormaPag & "'"
            RsTipoMonetario.FindFirst LcCriterioPes
            If Not RsTipoMonetario.NoMatch Then
               LcTipoMonetario = RsTipoMonetario("TPMONET")
            Else
               LcTipoMonetario = "03"
            End If
            LcValor = CCur(RsNota!ValorNota)
            'LcValor = Replace(LcValor, ",", ".")
            
            LcValorPago = CCur(RsNota!ValorNota)
           ' LcValorPago = Replace(LcValorPago, ",", ".")
            If LcNovo Then RsContasReceber.AddNew
            
            RsContasReceber!NF = RsNota!NUMNF & ""
            RsContasReceber!Cliente = RsNota!Cliente
            RsContasReceber!TPMONET = LcTipoMonetario
            RsContasReceber!valor = LcValor
            RsContasReceber!Data = Format(RsNota!DTEMIS, "dd/mm/yy")
            RsContasReceber!DTVENC = Format(RsNota!Vencimento1, "dd/mm/yy")
            'RsContasReceber!DTPAGTO = Format(txt(12).Text, "dd/mm/yy")
            RsContasReceber!VALPAGO = LcValor
            RsContasReceber!tipord = "R"
            RsContasReceber!acrescimo = 0
            RsContasReceber.Update
       
     
Else
     If IsDate(RsNota!vencimento2) Then
        LcNumeroContas = 2
     Else
        LcNumeroContas = 1
     End If
            For a = 1 To LcNumeroContas
                LcCriterioPes = "XTPMONET='" & RsNota!FormaPag & "'"
                RsTipoMonetario.FindFirst LcCriterioPes
                If Not RsTipoMonetario.NoMatch Then
                   LcTipoMonetario = RsTipoMonetario("TPMONET")
                Else
                   LcTipoMonetario = "03"
                End If
                LcValor = CCur(RsNota!ValorNota) / LcNumeroContas
                'LcValor = Replace(LcValor, ",", ".")
                
                LcValorPago = CCur(RsNota!ValorNota)
                'LcValorPago = Replace(LcValorPago, ",", ".")
                If LcNovo Then RsContasReceber.AddNew
            
                RsContasReceber!NF = RsNota!NUMNF & "/" & Right("00" & a, 2)
                RsContasReceber!Cliente = RsNota!Cliente
                RsContasReceber!TPMONET = LcTipoMonetario
                RsContasReceber!valor = LcValor
                RsContasReceber!Data = Format(RsNota!DTEMIS, "dd/mm/yy")
                Select Case a
                    Case Is = 1
                         If Not IsNull(RsNota!Vencimento1) Then
                            RsContasReceber("DTVENC") = CDate(Format(RsNota!Vencimento1, "dd/mm/yy"))
                         Else
                            RsContasReceber("DTVENC") = Format(RsNota!DTEMIS, "dd/mm/yy")
                         End If
                    Case Is = 2
                         RsContasReceber("DTVENC") = CDate(Format(RsNota!vencimento2, "dd/mm/yy"))
                    Case Is = 3
                         RsContasReceber("DTVENC") = CDate(Format(RsNota!vencimento3, "dd/mm/yy"))
                End Select
                'RsContasReceber!DTPAGTO = Format(txt(12).Text, "dd/mm/yy")
                RsContasReceber!VALPAGO = 0
                RsContasReceber!tipord = "R"
                RsContasReceber!acrescimo = 0
                RsContasReceber.Update
            Next
   
 End If
End Sub

