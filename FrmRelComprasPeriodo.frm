VERSION 5.00
Begin VB.Form FrmRelComprasPeriodo 
   BackColor       =   &H00D8C5B6&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Produtos Não comprados no Periodo"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6990
   Icon            =   "FrmRelComprasPeriodo.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   6990
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox fornecedor 
      Height          =   315
      Left            =   1440
      TabIndex        =   7
      Text            =   "fornecedor"
      Top             =   600
      Width           =   5175
   End
   Begin VB.CommandButton CmdGerar 
      Caption         =   "Gerar Relatorio"
      Height          =   615
      Left            =   1080
      TabIndex        =   6
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox Dias 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Text            =   "60"
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox Produto 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Fornecedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Previsao 
      BackStyle       =   0  'Transparent
      Caption         =   "Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Prevista"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Dias"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Produtos"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "FrmRelComprasPeriodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type TipoFor
      Codigo As String
      Nome As String
End Type
Private LcTamanho As Long
Private a As Integer
Private Mtfor() As TipoFor
Private Rel As New CrysCompraNaoRealizada

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
   Mtfor(LcTamanho).Nome = RsFornecedor!RAZAOSOC
   fornecedor.AddItem RsFornecedor!RAZAOSOC
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
Sub GeraRel()
On Error GoTo errGeracao
Dim Rs As ADODB.Recordset
Dim StrSql As String
StrSql = "Select * from relestoqueperiodo order by nome"
Set Rs = AbreRecordset(StrSql, True)
Load Relatorios
With Relatorios
     Rel.DiscardSavedData
     Rel.Database.SetDataSource Rs
     .CRViewer1.ReportSource = Rel
     setaformula
      .CRViewer1.ViewReport
End With
Relatorios.Show
Exit Sub
errGeracao:
MsgBox "Erro encontrado:" & err.Description
End Sub
Sub setaformula()
Dim lctitulo As String
lctitulo = "Produtos não Comprados apos " & Previsao.Caption & " - " & Dias.Text & " dias."
With Rel
For a = 1 To .FormulaFields.Count
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("Fone") Then .FormulaFields(a).Text = "totext('" & LcFone & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("EMPRESA") Then .FormulaFields(a).Text = "totext('" & LcEmpresa & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("ENDERECOEMPRESA") Then .FormulaFields(a).Text = "totext('" & LcEndereco & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = "TIPO" Then .FormulaFields(a).Text = "totext('" & Tipo & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("email") Then .FormulaFields(a).Text = "totext('" & Lcemail & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("Fornecedor") Then .FormulaFields(a).Text = "totext('" & fornecedor.Text & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("Titulo") Then
           .FormulaFields(a).Text = "totext('" & lctitulo & "')"
        End If
    Next
End With
End Sub
Private Sub CmdGerar_Click()
On Error Resume Next
Dim RsProduto   As ADODB.Recordset
Dim StrSql As String
Dim LcFormecedor As String
Dim LcCap As String
Dim LcTotal As Long
Dim StrWhere As String
StrSql = ""
LcCap = Me.Caption

If Not IsNumeric(Dias.Text) Then
   MsgBox "Informe quantos dias deve ser considerado para a verificação da Venda", 64, "AViso"
   Exit Sub
End If
Screen.MousePointer = vbHourglass

Dim LcAchou As Boolean
LcAchou = False
If Len(fornecedor.Text) > 0 Then
    For a = 0 To LcTamanho
      If Mtfor(a).Nome = fornecedor.Text Then
         LcFormecedor = Mtfor(a).Codigo
         LcAchou = True
         Exit For
      End If
    Next
Else
  LcFormecedor = ""
End If


StrSql = "Select codigo,nome,custo,(QuantEstoque/QtdMedida) as Estoque from produtos "
StrWhere = "where (ultimaAlteracao <'" & Format(CDate(Previsao.Caption), "yyyy-mm-dd") & "' or ultimaAlteracao is null) and desativado=0"
If Len(Produto.Text) > 0 Then
   StrWhere = StrWhere & " and nome like '%" & Produto.Text & "%'"
End If
If Len(LcFormecedor) > 0 Then
    StrWhere = StrWhere & " and Fornecedor='" & LcFormecedor & "'"
End If
If Len(StrWhere) > 0 Then StrSql = StrSql & StrWhere

StrSql = StrSql & " order by nome"
'Debug.Print StrSql
Set RsProduto = AbreRecordset(StrSql, True)
'==> Exclui os dados antigos
Me.Caption = "Excluindo dados de relatorios antigos"
DoEvents
StrSql = "Delete from relestoqueperiodo"
ExecutaSql StrSql
LcTotal = RsProduto.RecordCount

a = 1
Do Until RsProduto.EOF
   Dim RsVerificacao As ADODB.Recordset
   Me.Caption = "Verificando produto " & RsProduto!Nome & " " & a & " de " & LcTotal
   DoEvents
   '==> Verifica se Tem Lancamento
   StrSql = "Select codigo from itensentradanf where  data >='" & Format(CDate(Previsao.Caption), "yyyy-mm-dd") & "' and item='" & RsProduto!Codigo & "'  limit 1"
   'Debug.Print StrSql
   Set RsVerificacao = AbreRecordset(StrSql, True)
   If RsVerificacao.EOF Then
      '==> Nao tem o lancamento, entao vamos recuperar as informação
      Dim RsNota As ADODB.Recordset
      StrSql = "Select QTDE,data,NUMNF,fornecedor from itensentradanf where data <'" & Format(CDate(Previsao.Caption), "yyyy-mm-dd") & "' and item='" & RsProduto!Codigo & "' order by codigo desc limit 1"
      Set RsNota = AbreRecordset(StrSql, True)
      If Not RsNota.EOF Then
        Me.Caption = "Encontrado produto " & RsProduto!Nome
        DoEvents
         '==> Vamos Gravar na Tabela
         StrSql = "insert into relestoqueperiodo(CodigoProduto,Nome,Estoque,Custo,QuantUltimaCompra,DataUltimaCompra,NF,FORNECEDOR) Values("
         StrSql = StrSql & RsProduto!Codigo & ","
         StrSql = StrSql & "'" & Replace(RsProduto!Nome, "'", "''") & "',"
         StrSql = StrSql & Replace(RsProduto!Estoque, ",", ".") & ","
         StrSql = StrSql & Replace(RsProduto!Custo, ",", ".") & ","
         StrSql = StrSql & Replace(RsNota!QTDE, ",", ".") & ","
         StrSql = StrSql & "'" & Format(RsNota!Data, "yyyy-mm-dd") & "',"
         StrSql = StrSql & "'" & Replace(RsNota!NumNf, "'", "''") & "',"
         StrSql = StrSql & "'" & Replace(RsNota!fornecedor, "'", "''") & "')"
         ExecutaSql StrSql
      End If
     End If
     RsProduto.MoveNext
     a = a + 1
Loop
Me.Caption = "Montando o Relatorio..."
DoEvents
GeraRel
Me.Caption = LcCap
Screen.MousePointer = vbDefault
'MsgBox "Processo Terminado"
End Sub


Private Sub Dias_Change()
If IsNumeric(Dias.Text) Then
   Previsao.Caption = Format(Date - CInt(Dias.Text), "dd/mm/yy")
Else
   Previsao.Caption = ""
End If
End Sub

Private Sub Form_Load()
Carregaforn
If IsNumeric(Dias.Text) Then
   Previsao.Caption = Format(Date - CInt(Dias.Text), "dd/mm/yy")
Else
   Previsao.Caption = ""
End If
End Sub
