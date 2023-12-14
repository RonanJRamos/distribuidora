VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form RelatorioInventarioFiscal 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventario Fiscal"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   2835
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00808000&
      Caption         =   "Desativados"
      Height          =   1455
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   2535
      Begin VB.OptionButton DesativadosNao 
         BackColor       =   &H00808000&
         Caption         =   "Não Mostrar Desativados"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.OptionButton DesativadosSim 
         BackColor       =   &H00808000&
         Caption         =   "Somente Desativados"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   510
         Width           =   2295
      End
      Begin VB.OptionButton DesativadosTodos 
         BackColor       =   &H00808000&
         Caption         =   "Todos os Produtos"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   2295
      End
   End
   Begin MSMask.MaskEdBox DataI 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "99/99/99"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exibir Relatorio"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   2535
   End
   Begin MSMask.MaskEdBox DataF 
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   360
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "99/99/99"
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "RelatorioInventarioFiscal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Rs As ADODB.Recordset
Private Rel As New CrysRelInventarioFiscal
Private a As Integer
Function SetaImpressora() As String
On Error Resume Next
    Rel.PrinterSetup Relatorios.hWnd
    SetaImpressora = Rel.PrinterName
    Relatorios.SetFocus
End Function
Private Sub Command1_Click()

On Error Resume Next
Dim StrSql          As String
LcCap = Me.Caption
Me.Caption = "Aguarde, processando os dados..."
Screen.MousePointer = 11
GeraDados
Screen.MousePointer = 0
Me.Caption = LcCap
Screen.MousePointer = vbHourglass

StrSql = "Select * from relatorioinvfiscal order by nome"

Set Rs = AbreRecordset(StrSql, True)
Load Relatorios
With Relatorios
     Rel.DiscardSavedData
     Rel.Database.SetDataSource Rs
     .CRViewer1.ReportSource = Rel
    ' setaformula
      .CRViewer1.ViewReport
End With
Relatorios.Show
Screen.MousePointer = vbDefault
End Sub
Sub GeraDados()
On Error GoTo errGera
Dim RsEntrada As ADODB.Recordset
Dim RsSaida As ADODB.Recordset
Dim RsRel   As ADODB.Recordset
Dim StrSql As String
Dim Quantidade As Double
Dim SaldoAnterior As Double

Dim LcWhereDesativado As String

If DesativadosSim.Value Then
  LcWhereDesativado = " And produtos.Desativado<>0"
End If
If DesativadosNao.Value Then
  LcWhereDesativado = " And produtos.Desativado=0"
End If

'==> Busca a Entrada
StrSql = "SELECT Sum(itensentradanf.QTDE) AS SomaDeQTDE, itensentradanf.ITEM, produtos.NOME " & _
         "FROM (entradanf INNER JOIN itensentradanf ON entradanf.codigo = itensentradanf.CodigoNota) INNER JOIN produtos ON itensentradanf.ITEM = produtos.codigo " & _
         "WHERE (((entradanf.DATA) Between '" & Format(DataI.Text, "yyyy-mm-dd") & "' And '" & Format(DataF.Text, "yyyy-mm-dd") & "')) " & LcWhereDesativado & _
         " GROUP BY itensentradanf.ITEM, produtos.NOME;"
Debug.Print StrSql
Set RsEntrada = AbreRecordset(StrSql, True)

'==> Busca as saidas
StrSql = "SELECT produtos.codigo, produtos.NOME, alid052.QTDE, alid052.QTDUM, produtos.QtdMedida " & _
         "FROM (alid050 INNER JOIN alid052 ON alid050.NUMNF = alid052.NUMNF) INNER JOIN produtos ON alid052.codProd = produtos.codigo " & _
         "WHERE (((alid050.DTEMIS) Between '" & Format(DataI.Text, "yyyy-mm-dd") & "' And '" & Format(DataF.Text, "yyyy-mm-dd") & "')) " & LcWhereDesativado & _
         " GROUP BY produtos.codigo, produtos.NOME, alid052.QTDUM, produtos.QtdMedida " & _
         "ORDER BY produtos.codigo;"
Debug.Print StrSql
Set RsSaida = AbreRecordset(StrSql, True)

'==> Esclui a Tabela
afetados = ExecutaSql("Delete from relatorioinvfiscal")
'==> Inclui os dados da entrada
Debug.Print StrSql
Do Until RsEntrada.EOF
  SaldoAnterior = 0
  SaldoAnterior = BuscaSaldoAnteriorProduto(RsEntrada!Item)
  StrSql = "Insert into relatorioinvfiscal (CodProd,Nome,Entrada,Saidas,Anterior) vALUES ("
  StrSql = StrSql & RsEntrada!Item & ",'"
  StrSql = StrSql & Replace(RsEntrada!Nome, "'", "''") & "',"
  StrSql = StrSql & Replace(RsEntrada!SomaDeQTDE, ",", ".") & ","
  StrSql = StrSql & "0,"
  StrSql = StrSql & Replace(SaldoAnterior, ",", ".") & ")"
  afetados = ExecutaSql(StrSql)
 ' MsgBox DEscricaoErro
  'MsgBox StrSql
  RsEntrada.MoveNext
Loop
Do Until RsSaida.EOF
  '==> Verifica se ja existe o produto
  StrSql = "Select * from relatorioinvfiscal where CodProd=" & RsSaida!codigo
  Set RsRel = AbreRecordset(StrSql, True)
  'MsgBox DEscricaoErro
  Quantidade = 0
  Quantidade = (RsSaida!QTDE * RsSaida!QTDUM) / IIf(RsSaida!QtdMedida > 0, RsSaida!QtdMedida, 1)
  If Not RsRel.EOF Then
     StrSql = "Update relatorioinvfiscal set Saidas=" & Replace(CStr(RsRel!Saidas + Quantidade), ",", ".")
     StrSql = StrSql & " Where codigo=" & RsRel!codigo
  Else
     SaldoAnterior = 0
    SaldoAnterior = BuscaSaldoAnteriorProduto(RsSaida!codigo)

    StrSql = "Insert into relatorioinvfiscal (CodProd,Nome,Entrada,Saidas,Anterior) vALUES ("
    StrSql = StrSql & RsSaida!codigo & ",'"
    StrSql = StrSql & Replace(RsSaida!Nome, "'", "''") & "',"
    StrSql = StrSql & "0,"
    StrSql = StrSql & Replace(CStr(Quantidade), ",", ".") & ","
    StrSql = StrSql & Replace(SaldoAnterior, ",", ".") & ")"
  End If
  afetados = ExecutaSql(StrSql)
  'MsgBox StrSql
 ' MsgBox DEscricaoErro
  RsSaida.MoveNext
Loop

Exit Sub
errGera:
MsgBox err.Description & "  " & err.Number
Resume 0
End Sub
Function BuscaSaldoAnteriorProduto(codigoproduto As Long) As Double
Dim StrSql As String
Dim RsEntrada As ADODB.Recordset
Dim RsSaida As ADODB.Recordset
Dim SaldoAnterior As Double

StrSql = "SELECT Sum(itensentradanf.QTDE) AS somadeqtde, itensentradanf.ITEM, produtos.NOME " & _
         "FROM (entradanf INNER JOIN itensentradanf ON entradanf.codigo = itensentradanf.CodigoNota) INNER JOIN produtos ON itensentradanf.ITEM = produtos.codigo " & _
         "Where (((entradanf.Data) < #" & Format(DataI.Text, "mm/dd/yy") & "#)) " & _
         "GROUP BY itensentradanf.ITEM, produtos.NOME " & _
         "HAVING (((itensentradanf.ITEM)=" & codigoproduto & "));"
         
Set RsEntrada = AbreRecordset(StrSql, True)

'==> Busca as saidas
StrSql = "SELECT produtos.codigo, produtos.NOME, alid052.QTDE, alid052.QTDUM, produtos.QtdMedida " & _
         "FROM (alid050 INNER JOIN alid052 ON alid050.NUMNF = alid052.NUMNF) INNER JOIN produtos ON alid052.codProd = produtos.codigo " & _
         "Where (((produtos.Codigo) = " & codigoproduto & ") And ((alid050.DTEMIS) < #" & Format(DataI.Text, "mm/dd/yy") & "#)) " & _
         "ORDER BY produtos.codigo;"

         
Set RsSaida = AbreRecordset(StrSql, True)

SaldoAnterior = 0
If Not RsEntrada.EOF Then
   SaldoAnterior = RsEntrada!SomaDeQTDE
End If
Do Until RsSaida.EOF
   Quantidade = 0
   Quantidade = (RsSaida!QTDE * RsSaida!QTDUM) / IIf(RsSaida!QtdMedida > 0, RsSaida!QtdMedida, 1)
   SaldoAnterior = SaldoAnterior - Quantidade
   RsSaida.MoveNext
Loop
BuscaSaldoAnteriorProduto = SaldoAnterior
End Function

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
   LcFone = RsEmpresa!Fone & ""
   Lcemail = RsEmpresa!Email & ""
   
End If
Set RsEmpresa = Nothing
lctitulo = "Relatorio de Produtos (Estoque Galpao)"
With Rel
'Exit Sub
For a = 1 To .FormulaFields.Count
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("Fone") Then .FormulaFields(a).Text = "totext('" & LcFone & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("EMPRESA") Then .FormulaFields(a).Text = "totext('" & LcEmpresa & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("ENDERECOEMPRESA") Then .FormulaFields(a).Text = "totext('" & LcEndereco & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = "TIPO" Then .FormulaFields(a).Text = "totext('" & Tipo & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("email") Then .FormulaFields(a).Text = "totext('" & Lcemail & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("Titulo") Then
           .FormulaFields(a).Text = "totext('" & lctitulo & "')"
        End If
    Next
End With
End Sub

