VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form RelNotaSaidaResumoProdutos 
   BackColor       =   &H00DFCCA4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resumo dos Produtos de Saida"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4560
   Icon            =   "RelNotaSaidaResumoProdutos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox CodCliente 
      BackColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   3480
      TabIndex        =   15
      Top             =   2520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox Cliente 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   4335
   End
   Begin VB.TextBox Produto 
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   4335
   End
   Begin VB.TextBox CodVendedor 
      BackColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   3480
      TabIndex        =   11
      Top             =   1800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox Vendedor 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   4335
   End
   Begin VB.TextBox CodFornecedor 
      BackColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   3480
      TabIndex        =   10
      Top             =   1080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox Fornecedor 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   4335
   End
   Begin Crystal.CrystalReport Cryrelatorio 
      Left            =   3480
      Top             =   -120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.CommandButton CmdSair 
      BackColor       =   &H00CBC1AF&
      Caption         =   "Fechar"
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
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3720
      Width           =   2175
   End
   Begin VB.CommandButton cmdGerar 
      BackColor       =   &H00CBC1AF&
      Caption         =   "Gerar rel"
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3720
      Width           =   2175
   End
   Begin MSMask.MaskEdBox Dataf 
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   360
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
      TabIndex        =   0
      Top             =   360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Produto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Fornecedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo da Venda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "RelNotaSaidaResumoProdutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type dados_Resumo
    Codigo As String
    Nome As String
    Quantidade As Double
    CustoTotal As Double
    VendaTotal As Double
    Lucro      As Double
    Percentual As Double
End Type
Private Type Dados_for
    Codigo As String
    Nome As String
End Type
Private Type DadosV
    Codigo As String
    Nome As String
End Type
Private Type DadosC
    Codigo As String
    Nome As String
End Type
Private Mtfor() As Dados_for
Private MtV() As DadosV
Private MtC() As DadosC
Sub carregafor()
Dim Rs As Recordset
Dim db As Database
Dim a As Integer
'====carrega Fornecedor
Set db = OpenDatabase(GLBase)
Set Rs = db.OpenRecordset("Select * from alid002 order by razaosoc")
a = 0
ReDim Mtfor(0)
Do Until Rs.EOF
    DoEvents
   ReDim Preserve Mtfor(a)
   Mtfor(a).Nome = Rs!RazaoSoc & ""
   Mtfor(a).Codigo = Rs!Codigo & ""
   fornecedor.AddItem Rs!RazaoSoc & ""
   a = a + 1
  Rs.MoveNext
Loop
Rs.Close
Set Rs = Nothing
'====carrega vendedor
Set Rs = db.OpenRecordset("select * from alid200 order by Nome")
a = 0
ReDim MtV(0)
Do Until Rs.EOF
    DoEvents
   ReDim Preserve MtV(a)
   MtV(a).Nome = Rs!Nome & ""
   MtV(a).Codigo = Rs!Codigo
   Vendedor.AddItem Rs!Nome
   a = a + 1
  Rs.MoveNext
Loop
Rs.Close
Set Rs = Nothing
'====carrega Cliente
Set Rs = db.OpenRecordset("select CODIGO,RAZAOSOC from ALID001 order by razaosoc")
a = 0
ReDim MtC(0)
Do Until Rs.EOF
    DoEvents
   ReDim Preserve MtC(a)
   MtC(a).Nome = Rs!RazaoSoc & ""
   MtC(a).Codigo = Rs!Codigo
   Cliente.AddItem Rs!RazaoSoc
   a = a + 1
  Rs.MoveNext
Loop
Rs.Close
Set Rs = Nothing
db.Close
Set db = Nothing
End Sub
Private Sub CmdGerar_Click()
Dim LcFormula, LcCriterio As String, LcTipoSaida As Integer
Dim RsEmpresa As Recordset, RsOpcao As Recordset
Dim LcEmpresa, LcEndereco, LcFone, LcVer, LcCap, LcVer1 As String
Dim a As Integer
LcCap = Me.Caption
Me.Caption = "Aguarde, Gerando o Relatório..."
Screen.MousePointer = 11
CodFornecedor.Text = ""
a = 0
CodVendedor = ""
If Len(Vendedor.Text) > 0 Then
    For a = 0 To UBound(MtV)
       If MtV(a).Nome = Vendedor.Text Then
          CodVendedor.Text = MtV(a).Codigo
          Exit For
       End If
    Next
End If
a = 0
CodCliente.Text = ""
If Len(Cliente.Text) > 0 Then
    For a = 0 To UBound(MtC)
       If MtC(a).Nome = Cliente.Text Then
          CodCliente.Text = MtC(a).Codigo
          Exit For
       End If
    Next
End If
a = 0
CodFornecedor.Text = ""
If Len(fornecedor.Text) > 0 Then
    For a = 0 To UBound(Mtfor)
       If Mtfor(a).Nome = fornecedor.Text Then
          CodFornecedor.Text = Mtfor(a).Codigo
          Exit For
       End If
    Next
End If
GeraDados
CryRelatorio.DataFiles(0) = GLBase
CryRelatorio.ReportFileName = App.Path & "\ResumoProdutoVenda.rpt"
 
CryRelatorio.DiscardSavedData = True
CryRelatorio.WindowTop = 50
CryRelatorio.WindowWidth = 700
CryRelatorio.WindowLeft = 50
CryRelatorio.WindowHeight = 500
CryRelatorio.WindowTitle = "Resumo de Produtos Vendidos"

CryRelatorio.Formulas(0) = "titulo='Produtos Vendidos no Periodo de " & DataI.Text & " a " & DataF.Text & IIf(Len(Trim(fornecedor.Text)) > 0, " - " & Trim(fornecedor.Text), "") & "'"
 
LcTipoSaida = 0

CryRelatorio.Destination = LcTipoSaida
CryRelatorio.PrintReport
Me.Caption = LcCap
'RsOpcao.Close
'RsEmpresa.Close
'Dbbase.Close
Set RsOpcao = Nothing
Set RsEmpresa = Nothing
Set Dbbase = Nothing
Screen.MousePointer = 0
If CryRelatorio.LastErrorNumber > 0 Then MsgBox CryRelatorio.LastErrorString

End Sub

Private Sub CmdSair_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
carregafor
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
FrmPrincipal.SetFocus
End Sub

Sub GeraDados()
Dim RsNota      As ADODB.Recordset
Dim db          As Database
Dim RsProduto   As ADODB.Recordset
Dim StrSql      As String
Dim Mt()        As dados_Resumo
Dim a           As Integer
Dim primeiro    As Boolean
Dim b           As Integer
Dim CodUnidade  As Integer
Dim Quantidade  As Double
Dim QuantidadeUnidade As Double
Dim Achou       As Boolean
Dim Custo       As Double
Set db = OpenDatabase(GLBase)
If Len(CodFornecedor.Text) = 0 Then
    StrSql = "SELECT ALID050.DTEMIS, ALID050.CFOP, ALID050.STATUS, ALID052.QTDE, ALID052.UNIMED, ALID052.QTDUM, ALID052.descricao, ALID052.codProd, ALID052.VALUNIT, Produtos.Fornecedor " & _
           "FROM Produtos INNER JOIN (ALID050 INNER JOIN ALID052 ON ALID050.NUMNF = ALID052.NUMNF) ON Produtos.codigo = ALID052.codProd " & _
           " WHERE (((ALID050.DTEMIS) Between #" & Format(DataI.Text, "mm/dd/yy") & "# And #" & Format(DataF.Text, "mm/dd/yy") & "#) and ALID050.status Like 'autorizado%')"
Else
    StrSql = "SELECT alid050.DTEMIS, ALID050.CFOP, ALID050.STATUS, alid052.codProd, ALID052.QTDUM, alid052.ITEM, alid052.QTDE, alid052.VALUNIT, alid052.UNIMED, alid052.descricao, produtos.QuantEstoque, produtos.Fornecedor " & _
           "FROM (alid050 INNER JOIN alid052 ON alid050.NUMNF = alid052.NUMNF) INNER JOIN produtos ON alid052.codProd = produtos.codigo " & _
           " WHERE (((alid050.DTEMIS) Between #" & Format(DataI.Text, "mm/dd/yy") & "# And #" & Format(DataF.Text, "mm/dd/yy") & "#) and ALID050.status Like 'autorizado%' and ((produtos.Fornecedor)='" & CodFornecedor.Text & "'))"
End If

If Len(CodVendedor.Text) > 0 Then
    StrSql = StrSql & " and alid050.Vendedor='" & CLng(CodVendedor.Text) & "'"
End If
If Len(CodCliente.Text) > 0 Then
    StrSql = StrSql & " and (alid050.Cliente='" & CLng(CodCliente.Text) & "' or alid050.Cliente='" & CodCliente.Text & "')"
End If
StrSql = StrSql & " and alid052.Descricao like '" & produto.Text & "%'"
Debug.Print StrSql
Set RsNota = AbreRecordset(StrSql, True)
StrSql = "Select * from produtos order by codigo"
Debug.Print StrSql
Set RsProduto = AbreRecordset(StrSql, True)
'==>Exclui dados do banco
db.Execute ("delete from relResumoProdutoVenda")
primeiro = True
Do Until RsNota.EOF
  If RsNota!Status <> "CANCELADA" Then
        RsProduto.Filter = "codigo=" & RsNota!codProd
       ' If RsProduto!Codigo = 239 Then Stop
        If Not RsProduto.EOF Then
              'Custo = RsProduto!CustoTotal / IIf(Not IsNull(RsProduto!QtdMedida), IIf(RsProduto!QtdMedida > 0, RsProduto!QtdMedida, 1), 0)
              Custo = RsProduto!Custo / IIf(Not IsNull(RsProduto!QtdMedida), IIf(RsProduto!QtdMedida > 0, RsProduto!QtdMedida, 1), 0)
        Else
              Custo = 0
        End If
        If primeiro Then
          '==> Localiza o produto
           Quantidade = RsNota!Qtde * RsNota!QTDUM
           ReDim Mt(a)
           Mt(a).Codigo = RsNota!codProd
           Mt(a).CustoTotal = Custo * (RsNota!Qtde * RsNota!QTDUM)
           Mt(a).Nome = RsNota!Descricao
           Mt(a).Quantidade = Quantidade
           Mt(a).VendaTotal = RsNota!Qtde * RsNota!VALUNIT
           Mt(a).Lucro = Mt(a).VendaTotal - Mt(a).CustoTotal
           Mt(a).Percentual = (Mt(a).Lucro / Mt(a).VendaTotal) * 100
           a = a + 1
           primeiro = False
        Else
          '==> Ja tem lancamento, vamos veificar se tem o produto
          Achou = False
          For b = 0 To UBound(Mt)
              If Mt(b).Codigo = RsNota!codProd Then
                 Achou = True
                 Exit For
              End If
          Next
          If Achou Then
             Quantidade = RsNota!Qtde * RsNota!QTDUM
             Mt(b).CustoTotal = Mt(b).CustoTotal + (Custo * Quantidade)
             Mt(b).Quantidade = Mt(b).Quantidade + Quantidade
             Mt(b).VendaTotal = Mt(b).VendaTotal + (RsNota!Qtde * RsNota!VALUNIT)
             Mt(b).Lucro = Mt(b).VendaTotal - Mt(b).CustoTotal
             Mt(b).Percentual = (Mt(b).Lucro / Mt(b).VendaTotal) * 100
          Else
            Quantidade = RsNota!Qtde * RsNota!QTDUM
            ReDim Preserve Mt(a)
            Mt(a).Codigo = RsNota!codProd
            Mt(a).CustoTotal = Custo '* (RsNota!QTDE * RsNota!QTDUM)
            Mt(a).Nome = RsNota!Descricao
            Mt(a).Quantidade = Quantidade
            Mt(a).VendaTotal = RsNota!Qtde * RsNota!VALUNIT
            Mt(a).Lucro = Mt(a).VendaTotal - Mt(a).CustoTotal
            Mt(a).Percentual = IIf(Mt(a).VendaTotal > 0, (Mt(a).Lucro / IIf(Mt(a).VendaTotal > 0, Mt(a).VendaTotal, 1)) * 100, 0)
            a = a + 1
          End If
        End If
   End If
   RsNota.MoveNext
Loop

'==> Busca al

Set db = OpenDatabase(GLBase)
If Len(CodFornecedor.Text) = 0 Then
    StrSql = "SELECT saidas.DTEMIS, saidas.CFOP, saidas.STATUS, saidasdados.QTDE, saidasdados.UNIMED, saidasdados.QTDUM, saidasdados.descricao, saidasdados.codProd, saidasdados.VALUNIT, Produtos.Fornecedor " & _
           "FROM Produtos INNER JOIN (saidas INNER JOIN saidasdados ON saidas.NUMNF = saidasdados.NUMNF) ON Produtos.codigo = saidasdados.codProd " & _
           " WHERE (((saidas.DTEMIS) Between #" & Format(DataI.Text, "mm/dd/yy") & "# And #" & Format(DataF.Text, "mm/dd/yy") & "#))"
Else
    StrSql = "SELECT saidas.DTEMIS, saidas.CFOP, saidas.STATUS, saidasdados.codProd, saidasdados.QTDUM, saidasdados.ITEM, saidasdados.QTDE, saidasdados.VALUNIT, saidasdados.UNIMED, saidasdados.descricao, produtos.QuantEstoque, produtos.Fornecedor " & _
           "FROM (saidas INNER JOIN saidasdados ON saidas.NUMNF = saidasdados.NUMNF) INNER JOIN produtos ON saidasdados.codProd = produtos.codigo " & _
           " WHERE (((saidas.DTEMIS) Between #" & Format(DataI.Text, "mm/dd/yy") & "# And #" & Format(DataF.Text, "mm/dd/yy") & "#) AND ((produtos.Fornecedor)='" & CodFornecedor.Text & "'))"
End If

If Len(CodVendedor.Text) > 0 Then
    StrSql = StrSql & " and saidas.Vendedor='" & CodVendedor.Text & "'"
End If
If Len(CodCliente.Text) > 0 Then
    StrSql = StrSql & " and saidas.Cliente='" & CodCliente.Text & "'"
End If
StrSql = StrSql & " and saidasdados.Descricao like '" & produto.Text & "%'"
Set RsNota = AbreRecordset(StrSql, True)
StrSql = "Select * from produtos order by codigo"
Debug.Print StrSql
Set RsProduto = AbreRecordset(StrSql, True)
'==>Exclui dados do banco
db.Execute ("delete from relResumoProdutoVenda")
'primeiro = True
Do Until RsNota.EOF
  If RsNota!Status <> "CANCELADA" Then
        RsProduto.Filter = "codigo=" & RsNota!codProd
       ' If RsProduto!Codigo = 239 Then Stop
        If Not RsProduto.EOF Then
              'Custo = RsProduto!CustoTotal / IIf(Not IsNull(RsProduto!QtdMedida), IIf(RsProduto!QtdMedida > 0, RsProduto!QtdMedida, 1), 0)
              Custo = RsProduto!Custo / IIf(Not IsNull(RsProduto!QtdMedida), IIf(RsProduto!QtdMedida > 0, RsProduto!QtdMedida, 1), 0)
        Else
              Custo = 0
        End If
        If primeiro Then
          '==> Localiza o produto
           Quantidade = RsNota!Qtde * RsNota!QTDUM
           ReDim Mt(a)
           Mt(a).Codigo = RsNota!codProd
           Mt(a).CustoTotal = Custo * (RsNota!Qtde * RsNota!QTDUM)
           Mt(a).Nome = RsNota!Descricao
           Mt(a).Quantidade = Quantidade
           Mt(a).VendaTotal = RsNota!Qtde * RsNota!VALUNIT
           Mt(a).Lucro = Mt(a).VendaTotal - Mt(a).CustoTotal
           Mt(a).Percentual = (Mt(a).Lucro / Mt(a).VendaTotal) * 100
           a = a + 1
           primeiro = False
        Else
          '==> Ja tem lancamento, vamos veificar se tem o produto
          Achou = False
          For b = 0 To UBound(Mt)
              If Mt(b).Codigo = RsNota!codProd Then
                 Achou = True
                 Exit For
              End If
          Next
          If Achou Then
             Quantidade = RsNota!Qtde * RsNota!QTDUM
             Mt(b).CustoTotal = Mt(b).CustoTotal + (Custo * Quantidade)
             Mt(b).Quantidade = Mt(b).Quantidade + Quantidade
             Mt(b).VendaTotal = Mt(b).VendaTotal + (RsNota!Qtde * RsNota!VALUNIT)
             Mt(b).Lucro = Mt(b).VendaTotal - Mt(b).CustoTotal
             Mt(b).Percentual = (Mt(b).Lucro / Mt(b).VendaTotal) * 100
          Else
            Quantidade = RsNota!Qtde * RsNota!QTDUM
            ReDim Preserve Mt(a)
            Mt(a).Codigo = RsNota!codProd
            Mt(a).CustoTotal = Custo * (RsNota!Qtde * RsNota!QTDUM)
            Mt(a).Nome = RsNota!Descricao
            Mt(a).Quantidade = Quantidade
            Mt(a).VendaTotal = RsNota!Qtde * RsNota!VALUNIT
            Mt(a).Lucro = Mt(a).VendaTotal - Mt(a).CustoTotal
            Mt(a).Percentual = IIf(Mt(a).VendaTotal > 0, (Mt(a).Lucro / IIf(Mt(a).VendaTotal > 0, Mt(a).VendaTotal, 1)) * 100, 0)
            a = a + 1
          End If
        End If
   End If
   RsNota.MoveNext
Loop


'==> Grava os dados na tb
On Error Resume Next
err.Number = 0
a = UBound(Mt)
If err.Number <> 0 Then Exit Sub
For a = 0 To UBound(Mt)
    RsProduto.Filter = "codigo=" & Mt(a).Codigo
    
    If Not RsProduto.EOF Then
       QuantidadeUnidade = Mt(a).Quantidade Mod IIf(Not IsNull(RsProduto!QtdMedida), IIf(RsProduto!QtdMedida > 0, RsProduto!QtdMedida, 1), 1)
       Quantidade = Int(Mt(a).Quantidade / IIf(Not IsNull(RsProduto!QtdMedida), IIf(RsProduto!QtdMedida > 0, RsProduto!QtdMedida, 1), 1))
    Else
       Quantidade = Mt(a).Quantidade
       QuantidadeUnidade = 0
    End If

    StrSql = "Insert into relResumoProdutoVenda (Codigo,Nome,Quantidade,QuantUnidade,ValorCustoTotal,ValorVendaTotal,Lucro,Percentual) Values(" & _
           Mt(a).Codigo & ",'" & _
           Mt(a).Nome & "'," & _
           Replace(Quantidade, ",", ".") & "," & _
           Replace(QuantidadeUnidade, ",", ".") & "," & _
           Replace(Mt(a).CustoTotal, ",", ".") & "," & _
           Replace(Mt(a).VendaTotal, ",", ".") & "," & _
           Replace(Mt(a).Lucro, ",", ".") & "," & _
           Replace(Mt(a).Percentual, ",", ".") & ")"
    db.Execute StrSql
Next

End Sub

Private Sub fornecedor_Click()
Dim a As Integer
CodFornecedor.Text = ""
For a = 0 To UBound(Mtfor)
   If Mtfor(a).Nome = fornecedor.Text Then
      CodFornecedor.Text = Mtfor(a).Codigo
      Exit For
   End If
Next
End Sub

Private Sub Fornecedor_LostFocus()
Dim a As Integer
CodFornecedor.Text = ""
For a = 0 To UBound(Mtfor)
   If Mtfor(a).Nome = fornecedor.Text Then
      CodFornecedor.Text = Mtfor(a).Codigo
      Exit For
   End If
Next
End Sub

