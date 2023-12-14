VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FrmPesquisaProdutos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pesquisa Produtos"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11430
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   11430
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok F2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9720
      TabIndex        =   4
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar F10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9720
      TabIndex        =   5
      Top             =   720
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pesquisar"
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1695
      Begin VB.OptionButton Parte 
         Caption         =   "Qualquer Parte"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton Inicio 
         Caption         =   "Inicio"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MostraCliente 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   7223
      _Version        =   393216
      Rows            =   17
      Cols            =   11
      FixedCols       =   0
      BackColor       =   -2147483624
      FocusRect       =   2
      SelectionMode   =   1
      AllowUserResizing=   3
   End
   Begin VB.TextBox txt 
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   720
      Width           =   6975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Critério"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2040
      TabIndex        =   7
      Top             =   360
      Width           =   1020
   End
End
Attribute VB_Name = "FrmPesquisaProdutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LcCampo As String
Private a As Integer
Private Sub CmdCancelar_Click()
On Error Resume Next
Me.Visible = False
FrmLocacao.Txt(4).SetFocus
 
End Sub

Private Sub CmdCancelar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Me.Visible = False
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{C}"
End Sub

Private Sub CmdOk_Click()
ExibePesquisa
End Sub

Private Sub CmdOk_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Me.Visible = False
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{C}"
End Sub

Private Sub Form_Activate()
Me.Refresh
ExibePesquisa
MostraCliente.SetFocus
End Sub

Private Sub Form_Load()
On Error Resume Next
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
LcCampo = "Produto"
GeraGrid
ExibePesquisa
MostraCliente.SetFocus
End Sub
Function GeraGrid()
MostraCliente.ColAlignment(0) = 1
MostraCliente.ColAlignment(1) = 1
MostraCliente.ColAlignment(2) = 1
MostraCliente.ColAlignment(3) = 1
MostraCliente.ColAlignment(4) = 1
MostraCliente.ColAlignment(5) = 1
MostraCliente.ColAlignment(6) = 1
MostraCliente.ColAlignment(9) = 1
MostraCliente.ColAlignment(10) = 1
MostraCliente.ColWidth(0) = 900
MostraCliente.ColWidth(1) = 5000
MostraCliente.ColWidth(2) = 1200
MostraCliente.ColWidth(3) = 1000
MostraCliente.ColWidth(4) = 1000
MostraCliente.ColWidth(5) = 1000 ' 0
MostraCliente.ColWidth(6) = 0 ' 900
MostraCliente.ColWidth(7) = 1000 ' 0
MostraCliente.ColWidth(8) = 0 '900
MostraCliente.ColWidth(9) = 900
MostraCliente.ColWidth(10) = 900
MostraCliente.TextMatrix(0, 0) = "Código"
MostraCliente.TextMatrix(0, 1) = "Nome"
MostraCliente.TextMatrix(0, 2) = "Preço Venda"
MostraCliente.TextMatrix(0, 3) = "Preço Mim"
MostraCliente.TextMatrix(0, 4) = "Limite Pr."

MostraCliente.TextMatrix(0, 5) = "Embalagem"
MostraCliente.TextMatrix(0, 7) = "Estoque"
MostraCliente.TextMatrix(0, 9) = "Und Est."
MostraCliente.TextMatrix(0, 10) = "Seg."
LcTamanhoGrid = 1
End Function
Function ExibePesquisa()
On Error Resume Next
Dim RsProduto As ADODB.Recordset, RsHistorico As Recordset, RsUnidade As Recordset
Dim LcCriSql As String, LcCriterio As String, LcUnidade As String
Dim LcTamanho, a As Long
Dim Est As ControleDb
Set Est = New ControleDb
'Verifica se Selecionou todos
If Len(Txt.Text) = 0 Then
    If Len(GlCriterioSql) > 0 Then
       LcCriSql = GlCriterioSql
    Else
        If Len(Trim(Txt.Text)) = 0 Then
           msg = "Aguarde, Criando Lista de Produtos..."
           If GlFormA.Name = "FrmEntradaProduto" Then
               LcCriSql = "select * From produtos where nome like '%' order by  nome"
           Else
               LcCriSql = "select * From produtos where nome like '%' and Desativado=0 order by  nome"
           End If
        Else
          msg = "Aguarde, Filtrando Membros Começados com " & UCase(Txt.Text)
          If Inicio Then
             If GlFormA.Name = "FrmEntradaProduto" Then
                LcCriSql = "select * From produtos where nome like '" & UCase(Txt.Text) & "%' order by nome"
             Else
                LcCriSql = "select * From produtos where nome like '" & UCase(Txt.Text) & "%'  and Desativado=0 order by nome"
             End If
             
         Else
            If GlFormA.Name = "FrmEntradaProduto" Then
               LcCriSql = "select * From produtos where nome like '%" & UCase(Txt.Text) & "%' order by nome"
            Else
               LcCriSql = "select * From produtos where nome like '%" & UCase(Txt.Text) & "%'  and Desativado=0 order by nome"
            End If
            
          End If
        End If
    End If
Else
    If Len(Trim(Txt.Text)) = 0 Then
           msg = "Aguarde, Criando Lista de Produtos..."
           If GlFormA.Name = "FrmEntradaProduto" Then
               LcCriSql = "select * From produtos where nome like '%' order by  nome"
           Else
               LcCriSql = "select * From produtos where nome like '%' and Desativado=0 order by  nome"
           End If
          
        Else
          msg = "Aguarde, Filtrando Membros Começados com " & UCase(Txt.Text)
          If Inicio Then
             If GlFormA.Name = "FrmEntradaProduto" Then
                LcCriSql = "select * From produtos where nome like '" & UCase(Txt.Text) & "%' order by nome"
             Else
                LcCriSql = "select * From produtos where nome like '" & UCase(Txt.Text) & "%'  and Desativado=0 order by nome"
             End If
             
         Else
            If GlFormA.Name = "FrmEntradaProduto" Then
               LcCriSql = "select * From produtos where nome like '%" & UCase(Txt.Text) & "%' order by nome"
            Else
               LcCriSql = "select * From produtos where nome like '%" & UCase(Txt.Text) & "%'  and Desativado=0 order by nome"
            End If
            
          End If
      End If
End If

'Set DbBase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
AbreBase
'abreconexao
Set RsProduto = AbreRecordset(LcCriSql, True) ', dbOpenDynaset, dbSeeChanges, dbOptimistic) ', dbOpenDynaset)
Set RsUnidade = Dbbase.OpenRecordset("select * From alid004", dbOpenDynaset, dbSeeChanges, dbOptimistic)  ', dbOpenDynaset)
'MsgBox LcCriSql
'Set RsHistorico = DbBase.OpenRecordset("HistoricoLocacao", dbOpenTable)
'RsHistorico.Index = "Produto"

LcTamanho = MostraCliente.Rows
a = 2
Me.Caption = msg
MostraCliente.Rows = 1

Do Until RsProduto.EOF
 If err.Number <> 0 Then
        
    If err <> 94 Then
        Exit Do
    Else
       err = 0
    End If
 End If
  If Len(Trim(RsProduto!Nome)) > 0 Then
   If Not IsNull(RsProduto!Nome) Then
     LcCriterio = "Cod='" & RsProduto!UnidMedida & "'"
     RsUnidade.FindFirst LcCriterio
     If Not RsUnidade.NoMatch Then
        LcUnidade = RsUnidade!Simbolo
     End If
     Est.CodProduto = RsProduto!Codigo
     MostraCliente.Rows = a
     MostraCliente.TextMatrix(a - 1, 0) = RsProduto!Codigo & ""
     MostraCliente.TextMatrix(a - 1, 1) = RsProduto!Nome & ""
     MostraCliente.TextMatrix(a - 1, 2) = Format(RsProduto!Preco & "", "currency")
     MostraCliente.TextMatrix(a - 1, 3) = Format(Est.PrecoMinimo & "", "currency")
     MostraCliente.TextMatrix(a - 1, 4) = Format(Est.LimiteVenda & "", "currency")
     MostraCliente.TextMatrix(a - 1, 5) = LcUnidade & " C/ " & RsProduto!QtdMedida
     MostraCliente.TextMatrix(a - 1, 6) = Format(RsProduto!Preco & "", "currency")
     MostraCliente.TextMatrix(a - 1, 7) = Est.EstoqueTotalFechado & ""
     MostraCliente.TextMatrix(a - 1, 8) = RsProduto!QtdMedida & ""
     MostraCliente.TextMatrix(a - 1, 9) = Est.EstoqueTotalUnitario & ""
     MostraCliente.TextMatrix(a - 1, 10) = Est.EstoqueSegurancaTotalFechado & ""
     '===> Muda a Cor dos Itens que estiverem com o Estoque zero
     If (CDbl(Est.EstoqueTotalFechado) + CDbl(Est.EstoqueTotalUnitario)) <= 0 Then
         Cor = &HC000&
         MostraCliente.Row = a - 1
          For x = 0 To 9
                MostraCliente.Col = x
                MostraCliente.CellBackColor = Cor
          Next
     End If
     If GlFormA.Name = "FrmEntradaProduto" Then
        If RsProduto!Desativado Then
            Cor = &H8080FF
            MostraCliente.Row = a - 1
             For x = 0 To 9
                   MostraCliente.Col = x
                   MostraCliente.CellBackColor = Cor
             Next
        End If
     End If
     a = a + 1
     'RsProduto.MoveNext
    End If
  
  End If
 RsProduto.MoveNext
Loop
Set Est = Nothing
LcCriSql = ""
Me.Caption = "Produtos Começados com " & Txt.Text
MostraCliente.SetFocus
If err <> 0 Then Txt.SetFocus
MostraCliente.Col = 1
MostraCliente.Row = 1
RsProduto.Close
RsUnidade.Close
Set RsUnidade = Nothing
Set RsProduto = Nothing

End Function

Private Sub Inicio_Click()
Txt.SetFocus
End Sub


Private Sub Inicio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Me.Visible = False
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{C}"
End Sub

Private Sub MostraCliente_DblClick()
On Error Resume Next
Dim a As Integer
Dim bb As Database
Dim LcCriSql As String

Dim RsProduto As ADODB.Recordset, RsUnidade As Recordset
LcSql1 = "Select * from alid004"
a = MostraCliente.Row

Set bb = OpenDatabase(GLBase, False, False) ' "dBASE III;")
Set RsUnidade = bb.OpenRecordset(LcSql1, dbOpenDynaset, dbSeeChanges, dbOptimistic) ', dbOpenDynaset)
PrecoVendaNormal = 0
If GlContrato And (GlFormA.Name = "FrmSaidaProduto" Or GlFormA.Name = "FrmProposta") Then
    LcSql = "SELECT ContratoDados.Valor,ContratoDados.CodProduto,ContratoFornecimento.DataI,ContratoFornecimento.DataF"
    LcSql = LcSql & " FROM ContratoDados INNER JOIN ContratoFornecimento ON ContratoDados.CodContrato = ContratoFornecimento.Codigo"
    If GlFormA.Name = "FrmSaidaProduto" Then
        LcSql = LcSql & " WHERE ContratoFornecimento.Cliente='" & FrmSaidaProduto.Txt(9).Text & "' and ContratoDados.CodProduto='" & MostraCliente.TextMatrix(a, 0) & "'"
    Else
        LcSql = LcSql & " WHERE ContratoFornecimento.Cliente='" & FrmProposta.Txt(9).Text & "' and ContratoDados.CodProduto='" & MostraCliente.TextMatrix(a, 0) & "'"
    End If
    LcSql = LcSql & " and ContratoFornecimento.DataI<#" & Format(Date, "mm/dd/yy") & "# and ContratoFornecimento.DataF>#" & Format(Date, "mm/dd/yy") & "#"
    Set RsProduto = AbreRecordset(LcSql, True) ', dbOpenDynaset, dbSeeChanges, dbOptimistic) ', dbOpenDynaset)
    Do Until RsProduto.EOF
        
        
        If RsProduto!CodProduto = MostraCliente.TextMatrix(a, 0) Then
            LcAchou = 1
            PrecoVendaNormal = CDbl(RsProduto!valor)
        
        End If
        RsProduto.MoveNext
    Loop
    RsProduto.Close
    Set RsProduto = Nothing
End If
LcCriSql = "select * from produtos where codigo=" & MostraCliente.TextMatrix(a, 0)
Set RsProduto = AbreRecordset(LcCriSql, True) ', dbOpenDynaset, dbSeeChanges, dbOptimistic) ', dbOpenDynaset)
 
 Select Case GlFormA.Name
 Case Is = "fichadeestoque"
           fichadeestoque.Codigo.Text = MostraCliente.TextMatrix(a, 0)
           fichadeestoque.Nome.Text = MostraCliente.TextMatrix(a, 1)
 Case Is = "ContratoFornecimento"
           ContratoFornecimento.CodProduto.Text = MostraCliente.TextMatrix(a, 0)
           ContratoFornecimento.Produto.Text = MostraCliente.TextMatrix(a, 1)
           ContratoFornecimento.ValorUnit.SetFocus
 Case Is = "FrmEntradaProduto"
        FrmEntradaProduto.Txt(2).Text = MostraCliente.TextMatrix(a, 1)
        FrmEntradaProduto.Txt(1).Text = MostraCliente.TextMatrix(a, 0)
        FrmEntradaProduto.Txt(4).Text = MostraCliente.TextMatrix(a, 8)
        LccriterioUn = "cod='" & RsProduto!UnidMedida & "'"
        RsUnidade.FindFirst LccriterioUn
        If Not RsUnidade.NoMatch Then
              FrmEntradaProduto.Unidade.Text = RsUnidade!Simbolo
        End If
        
        If MostraCliente.CellBackColor = &H8080FF Then
           FrmEntradaProduto.Label11.Caption = "PRODUTO DESATIVADO"
        Else
           FrmEntradaProduto.Label11.Caption = ""
        End If
        FrmEntradaProduto.Txt(3).SetFocus
        'RsProduto.FindFirst "cod='" & MostraCliente.TextMatrix(a - 1, 0) & "'"
        'FrmSaidaProduto.Valor(0).Text = MostraCliente.TextMatrix(a, 2)
        'If Not IsNull(RsProduto!Ptab) Then PrecoVendaNormal = RsProduto!Ptab / RsProduto!QTDUNIMED Else PrecoVendaNormal = 0
        'ComNormal = RsProduto!QTDUNIMED
        ' FrmSaidaProduto.minimo.Text = RsProduto!MPVENDA & ""
        '  If Not IsNull(RsProduto!MPVENDA) Then PrecoMimimodeVendaAlterado = RsProduto!MPVENDA / RsProduto!QTDUNIMED Else PrecoMimimodeVendaAlterado = 0
       ' FrmSaidaProduto.txt(5).Text = MostraCliente.TextMatrix(a, 5)
    Case Is = "FrmSaidaProdutoAlternativo"
        FrmSaidaProdutoAlternativo.Txt(2).Text = MostraCliente.TextMatrix(a, 1)
        FrmSaidaProdutoAlternativo.Txt(1).Text = MostraCliente.TextMatrix(a, 0)
        FrmSaidaProdutoAlternativo.Txt(4).Text = MostraCliente.TextMatrix(a, 8)
        If LcAchou = 0 Then
            FrmSaidaProdutoAlternativo.valor(0).Text = MostraCliente.TextMatrix(a, 2)
            If Not IsNull(RsProduto!Preco) Then PrecoVendaNormal = CCur(MostraCliente.TextMatrix(a, 2)) / CCur(MostraCliente.TextMatrix(a, 8)) Else PrecoVendaNormal = 0
        Else
            FrmSaidaProdutoAlternativo.valor(0).Text = AcertaNumero(CDbl(PrecoVendaNormal), 2)
        End If
        If Not RsProduto.EOF Then
            If LcAchou = 0 Then
                If Not IsNull(RsProduto!Preco) Then PrecoVendaNormal = CCur(MostraCliente.TextMatrix(a, 2)) / CCur(MostraCliente.TextMatrix(a, 8)) Else PrecoVendaNormal = 0
            End If
           ComNormal = CCur(MostraCliente.TextMatrix(a, 8))
           'LcCriterio = "codigo=" & MostraCliente.TextMatrix(a, 0)
           'RsProduto.FindFirst LcCriterio
           LccriterioUn = "cod='" & RsProduto!UnidMedida & "'"
           RsUnidade.FindFirst LccriterioUn
           If Not RsUnidade.NoMatch Then
              FrmSaidaProdutoAlternativo.Unidade.Text = RsUnidade!Simbolo
           End If
           FrmSaidaProdutoAlternativo.minimo.Text = RsProduto!MinimoVenda & ""
           FrmSaidaProdutoAlternativo.cst.Text = RsProduto!cst
           If Not IsNull(RsProduto!MinimoVenda) Then PrecoMimimodeVendaAlterado = CCur(MostraCliente.TextMatrix(a, 2)) / CCur(MostraCliente.TextMatrix(a, 8)) Else PrecoMimimodeVendaAlterado = 0
        End If
         If RsProduto!icms = 0 Or IsNull(RsProduto!icms) Then
            If Val(FrmSaidaProdutoAlternativo.cst.Text) = 60 Or Val(FrmSaidaProdutoAlternativo.cst.Text) = 160 Or Val(FrmSaidaProdutoAlternativo.cst.Text) = 260 Then
                    FrmSaidaProdutoAlternativo.icms.Text = "00"
            Else
              If Len(FrmSaidaProdutoAlternativo.Txt(5).Text) = 0 Then
                    FrmSaidaProdutoAlternativo.icms.Text = "18"
              Else
                    FrmSaidaProdutoAlternativo.icms = FrmSaidaProdutoAlternativo.Txt(5).Text
              End If
             End If
         Else
            FrmSaidaProdutoAlternativo.icms = RsProduto!icms
         End If


    Case Is = "FrmSaidaProduto"
        'AbreRecordset
        'LcCriSql = "select * from alid009 where codigo=" & MostraCliente.TextMatrix(a - 1, 0)
       ' Set RsProduto = AbreRecordsetLeitura(LcCriSql) ', dbOpenDynaset, dbSeeChanges, dbOptimistic) ', dbOpenDynaset)

        FrmSaidaProduto.Txt(2).Text = MostraCliente.TextMatrix(a, 1)
        FrmSaidaProduto.Txt(1).Text = MostraCliente.TextMatrix(a, 0)
        FrmSaidaProduto.Txt(4).Text = MostraCliente.TextMatrix(a, 8)
        If LcAchou = 0 Then
            FrmSaidaProduto.valor(0).Text = MostraCliente.TextMatrix(a, 2)
            If Not IsNull(RsProduto!Preco) Then PrecoVendaNormal = CCur(MostraCliente.TextMatrix(a, 2)) / CCur(MostraCliente.TextMatrix(a, 8)) Else PrecoVendaNormal = 0
        Else
            FrmSaidaProduto.valor(0).Text = AcertaNumero(CDbl(PrecoVendaNormal), 2)
        End If
        If Not RsProduto.EOF Then
            If LcAchou = 0 Then
                If Not IsNull(RsProduto!Preco) Then PrecoVendaNormal = CCur(MostraCliente.TextMatrix(a, 2)) / CCur(MostraCliente.TextMatrix(a, 8)) Else PrecoVendaNormal = 0
            End If
           ComNormal = CCur(MostraCliente.TextMatrix(a, 8))
           'LcCriterio = "codigo=" & MostraCliente.TextMatrix(a, 0)
           'RsProduto.FindFirst LcCriterio
           LccriterioUn = "cod='" & RsProduto!UnidMedida & "'"
           RsUnidade.FindFirst LccriterioUn
           If Not RsUnidade.NoMatch Then
              FrmSaidaProduto.Unidade.Text = RsUnidade!Simbolo
           End If
           FrmSaidaProduto.minimo.Text = RsProduto!MinimoVenda & ""
           FrmSaidaProduto.cst.Text = RsProduto!cst
           If Not IsNull(RsProduto!MinimoVenda) Then PrecoMimimodeVendaAlterado = CCur(MostraCliente.TextMatrix(a, 2)) / CCur(MostraCliente.TextMatrix(a, 8)) Else PrecoMimimodeVendaAlterado = 0
        End If
         If RsProduto!icms = 0 Or IsNull(RsProduto!icms) Then
            If Val(FrmSaidaProduto.cst.Text) = 60 Or Val(FrmSaidaProduto.cst.Text) = 160 Or Val(FrmSaidaProduto.cst.Text) = 260 Then
                    FrmSaidaProduto.icms.Text = "00"
            Else
              If Len(FrmSaidaProduto.Txt(5).Text) = 0 Then
                    FrmSaidaProduto.icms.Text = "18"
              Else
                    FrmSaidaProduto.icms = FrmSaidaProduto.Txt(5).Text
              End If
             End If
         Else
            FrmSaidaProduto.icms = RsProduto!icms
         End If

         
       ' FrmSaidaProduto.txt(5).Text = MostraCliente.TextMatrix(a, 5)
      Case Is = "FrmVales"
        'LcCriSql = "select * from alid009 where codigo=" & MostraCliente.TextMatrix(a - 1, 0)
       ' Set RsProduto = AbreRecordsetLeitura(LcCriSql) ', dbOpenDynaset, dbSeeChanges, dbOptimistic) ', dbOpenDynaset)

        FrmVales.Txt(2).Text = MostraCliente.TextMatrix(a, 1)
        FrmVales.Txt(1).Text = MostraCliente.TextMatrix(a, 0)
        FrmVales.Txt(4).Text = MostraCliente.TextMatrix(a, 8)
       ' RsProduto.FindFirst "cod='" & MostraCliente.TextMatrix(a - 1, 0) & "'"
        FrmVales.valor(0).Text = MostraCliente.TextMatrix(a, 2)
        If Not IsNull(RsProduto!Preco) Then PrecoVendaNormal = CCur(MostraCliente.TextMatrix(a, 2)) / CCur(MostraCliente.TextMatrix(a, 8)) Else PrecoVendaNormal = 0
        ComNormal = CCur(MostraCliente.TextMatrix(a, 8))
        'LcCriterio = "codigo=" & MostraCliente.TextMatrix(a, 0)
        'RsProduto.Find LcCriterio
        If Not RsProduto.EOF Then
           LccriterioUn = "cod='" & RsProduto!UnidMedida & "'"
           RsUnidade.FindFirst LccriterioUn
           If Not RsUnidade.NoMatch Then
              FrmVales.Unidade.Text = RsUnidade!Simbolo
           End If
        End If
        FrmVales.minimo.Text = RsProduto!MinimoVenda & ""
         FrmVales.cst.Text = RsProduto!cst
        If RsProduto!icms = 0 Or IsNull(RsProduto!icms) Then
            If Val(FrmVales.cst.Text) = 60 Or Val(FrmVales.cst.Text) = 160 Or Val(FrmVales.cst.Text) = 260 Then
                    FrmVales.icms.Text = "00"
            Else
              If Len(FrmVales.Txt(5).Text) = 0 Then
                    FrmVales.icms.Text = "18"
              Else
                    FrmVales.icms = FrmVales.Txt(5).Text
              End If
             End If
         Else
            FrmVales.icms = RsProduto!icms
         End If
         
         If Not IsNull(RsProduto!MinimoVenda) Then PrecoMimimodeVendaAlterado = CCur(MostraCliente.TextMatrix(a, 2)) / CCur(MostraCliente.TextMatrix(a, 8)) Else PrecoMimimodeVendaAlterado = 0
       ' FrmSaidaProduto.txt(5).Text = MostraCliente.TextMatrix(a, 5)
  
    Case Is = "FrmReajustaPreco"
        FrmReajustaPreco.bo.Text = MostraCliente.TextMatrix(a, 1)
        FrmReajustaPreco.Codigo.Text = MostraCliente.TextMatrix(a, 0)
    Case Is = "FrmPedido"
        FrmPedido.Txt(2).Text = MostraCliente.TextMatrix(a, 1)
        FrmPedido.Txt(1).Text = MostraCliente.TextMatrix(a, 0)
        FrmPedido.Txt(4).Text = MostraCliente.TextMatrix(a, 8)
       ' RsProduto.FindFirst "cod='" & MostraCliente.TextMatrix(a - 1, 0) & "'"
        'FrmPedido.Valor(0).Text = MostraCliente.TextMatrix(a, 2)
         ComNormal = CCur(MostraCliente.TextMatrix(a, 8))
         FrmPedido.minimo.Text = RsProduto!MinimoVenda & ""
         If Not IsNull(RsProduto!MinimoVenda) Then PrecoMimimodeVendaAlterado = RsProduto!MPVENDA / CCur(MostraCliente.TextMatrix(a, 8)) Else PrecoMimimodeVendaAlterado = 0
       ' FrmSaidaProduto.txt(5).Text = MostraCliente.TextMatrix(a, 5)
        'FrmSaidaProduto.txt(5).Text = MostraC
        
    Case Is = "Orcamento"

        'lcpesqunidade = "cod='" & MostraCliente.TextMatrix(a, 0) & "'"
        'RsProduto.Find lcpesqunidade
        If Not RsProduto.EOF Then
           orcamento.codigoproduto.Text = RsProduto!Codigo
           orcamento.NomeProduto.Text = RsProduto!Nome
           orcamento.ipi.Text = RsProduto!ipi & ""
           lcprocuraunidade = "cod='" & RsProduto!UnidMedida & "'"
           orcamento.Industria.Text = RsProduto!fornecedor & ""
           RsUnidade.FindFirst lcprocuraunidade
           If Not RsUnidade.NoMatch Then
              orcamento.Unidade.Text = RsUnidade!Simbolo
              orcamento.codigounidade.Text = RsUnidade!cod
           End If
           If Len(RsProduto!Preco) > 0 Then
              orcamento.Unitario.Text = RsProduto!Preco
              orcamento.preconormal.Text = RsProduto!Preco
           Else
              orcamento.Unitario.Text = 0
           End If
        End If
        orcamento.Unidade.SetFocus
        RsProduto.Close
        RsUnidade.Close
     Case Is = "FrmProposta"
        FrmProposta.Txt(2).Text = MostraCliente.TextMatrix(a, 1)
        FrmProposta.Txt(1).Text = MostraCliente.TextMatrix(a, 0)
        FrmProposta.Txt(4).Text = MostraCliente.TextMatrix(a, 8)
        If LcAchou = 0 Then
            FrmProposta.valor(0).Text = MostraCliente.TextMatrix(a, 2)
            If Not IsNull(RsProduto!Preco) Then PrecoVendaNormal = CCur(MostraCliente.TextMatrix(a, 2)) / CCur(MostraCliente.TextMatrix(a, 8)) Else PrecoVendaNormal = 0
        Else
            FrmProposta.valor(0).Text = AcertaNumero(CDbl(PrecoVendaNormal), 2)
        End If
        ComNormal = CCur(MostraCliente.TextMatrix(a, 8))
        'LcCriterio = "cod='" & MostraCliente.TextMatrix(a, 0) & "'"
        'RsProduto.FindFirst LcCriterio
        If Not RsProduto.EOF Then
           LccriterioUn = "cod='" & RsProduto!UnidMedida & "'"
           RsUnidade.FindFirst LccriterioUn
           If Not RsUnidade.NoMatch Then
              FrmProposta.Unidade.Text = RsUnidade!Simbolo
           End If
        End If
        FrmProposta.minimo.Text = RsProduto!MinimoVenda & ""
        FrmProposta.cst.Text = RsProduto!cst
         
        If RsProduto!icms = 0 Or IsNull(RsProduto!icms) Then
            If Val(FrmProposta.cst.Text) = 60 Or Val(FrmProposta.cst.Text) = 160 Or Val(FrmProposta.cst.Text) = 260 Then
                    FrmProposta.icms.Text = "00"
            Else
              If Len(FrmProposta.Txt(5).Text) = 0 Then
                    FrmProposta.icms.Text = "18"
              Else
                    FrmProposta.icms = FrmVales.Txt(5).Text
              End If
             End If
         Else
         
            FrmProposta.icms = RsProduto!icms
         End If
         If Not IsNull(RsProduto!MinimoVenda) Then PrecoMimimodeVendaAlterado = CCur(MostraCliente.TextMatrix(a, 2)) / CCur(MostraCliente.TextMatrix(a, 8)) Else PrecoMimimodeVendaAlterado = 0
       ' FrmSaidaProduto.txt(5).Text = MostraCliente.TextMatrix(a, 5)
        FrmProposta.Custo.Text = RsProduto!CustoTotal & ""
        If Not IsNumeric(FrmProposta.Custo.Text) Then FrmProposta.Custo.Text = 0
        
    End Select
    
 Me.Visible = False
 GlFormA.SetFocus
End Sub

Private Sub MostraCliente_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Dim RsProduto As ADODB.Recordset, RsUnidade As Recordset
Dim a As Long
Dim LcCriSql As String
If KeyCode = 13 Then
    MostraCliente_DblClick
    Exit Sub
End If

LcCriSql = "select * from alid009"
LcSql1 = "Select * from alid004"
AbreBase
'Set RsProduto = Dbbase.OpenRecordset(LcCriSql, dbOpenDynaset, dbSeeChanges, dbOptimistic) ', dbOpenDynaset)
Set RsUnidade = Dbbase.OpenRecordset(LcSql1, dbOpenDynaset, dbSeeChanges, dbOptimistic) ', dbOpenDynaset)

If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{C}"
If KeyCode = 13 Then
    a = MostraCliente.Row
    'abreconexao
    LcCriSql = "select * from produtos where codigo=" & MostraCliente.TextMatrix(a, 0)
    Set RsProduto = AbreRecordset(LcCriSql, True) ', dbOpenDynaset, dbSeeChanges, dbOptimistic) ', dbOpenDynaset)

    
     Select Case GlFormA.Name
 Case Is = "fichadeestoque"
           fichadeestoque.Codigo.Text = MostraCliente.TextMatrix(a, 0)
           fichadeestoque.Nome.Text = MostraCliente.TextMatrix(a, 1)

 Case Is = "FrmEntradaProduto"
        FrmEntradaProduto.Txt(2).Text = MostraCliente.TextMatrix(a, 1)
        FrmEntradaProduto.Txt(1).Text = MostraCliente.TextMatrix(a, 0)
        FrmEntradaProduto.Txt(4).Text = MostraCliente.TextMatrix(a, 8)
        LccriterioUn = "cod='" & RsProduto!UnidMedida & "'"
        RsUnidade.FindFirst LccriterioUn
        If Not RsUnidade.NoMatch Then
              FrmEntradaProduto.Unidade.Text = RsUnidade!Simbolo
        End If
         If MostraCliente.CellBackColor = &H8080FF Then
           FrmEntradaProduto.Label11.Caption = "PRODUTO DESATIVADO"
        Else
           FrmEntradaProduto.Label11.Caption = ""
        End If
        FrmEntradaProduto.Txt(3).SetFocus
        'RsProduto.FindFirst "cod='" & MostraCliente.TextMatrix(a - 1, 0) & "'"
        'FrmSaidaProduto.Valor(0).Text = MostraCliente.TextMatrix(a, 2)
        'If Not IsNull(RsProduto!Ptab) Then PrecoVendaNormal = RsProduto!Ptab / RsProduto!QTDUNIMED Else PrecoVendaNormal = 0
        'ComNormal = RsProduto!QTDUNIMED
        ' FrmSaidaProduto.minimo.Text = RsProduto!MPVENDA & ""
        '  If Not IsNull(RsProduto!MPVENDA) Then PrecoMimimodeVendaAlterado = RsProduto!MPVENDA / RsProduto!QTDUNIMED Else PrecoMimimodeVendaAlterado = 0
       ' FrmSaidaProduto.txt(5).Text = MostraCliente.TextMatrix(a, 5)
    Case Is = "FrmSaidaProduto"
        'AbreRecordset
        'LcCriSql = "select * from alid009 where codigo=" & MostraCliente.TextMatrix(a - 1, 0)
       ' Set RsProduto = AbreRecordsetLeitura(LcCriSql) ', dbOpenDynaset, dbSeeChanges, dbOptimistic) ', dbOpenDynaset)

        FrmSaidaProduto.Txt(2).Text = MostraCliente.TextMatrix(a, 1)
        FrmSaidaProduto.Txt(1).Text = MostraCliente.TextMatrix(a, 0)
        FrmSaidaProduto.Txt(4).Text = MostraCliente.TextMatrix(a, 8)
        'RsProduto.FindFirst "codigo=" & MostraCliente.TextMatrix(a - 1, 0)
        FrmSaidaProduto.valor(0).Text = MostraCliente.TextMatrix(a, 2)
        If Not RsProduto.EOF Then
           If Not IsNull(RsProduto!Preco) Then PrecoVendaNormal = CCur(MostraCliente.TextMatrix(a, 2)) / CCur(MostraCliente.TextMatrix(a, 8)) Else PrecoVendaNormal = 0
           ComNormal = CCur(MostraCliente.TextMatrix(a, 8))
           'LcCriterio = "codigo=" & MostraCliente.TextMatrix(a, 0)
           'RsProduto.FindFirst LcCriterio
           LccriterioUn = "cod='" & RsProduto!UnidMedida & "'"
           RsUnidade.FindFirst LccriterioUn
           If Not RsUnidade.NoMatch Then
              FrmSaidaProduto.Unidade.Text = RsUnidade!Simbolo
           End If
           FrmSaidaProduto.minimo.Text = RsProduto!MinimoVenda & ""
           FrmSaidaProduto.cst.Text = RsProduto!cst
           If Not IsNull(RsProduto!MinimoVenda) Then PrecoMimimodeVendaAlterado = CCur(MostraCliente.TextMatrix(a, 2)) / CCur(MostraCliente.TextMatrix(a, 8)) Else PrecoMimimodeVendaAlterado = 0
        End If
         If RsProduto!icms = 0 Or IsNull(RsProduto!icms) Then
            If Val(FrmSaidaProduto.cst.Text) = 60 Or Val(FrmSaidaProduto.cst.Text) = 160 Or Val(FrmSaidaProduto.cst.Text) = 260 Then
                    FrmSaidaProduto.icms.Text = "00"
            Else
              If Len(FrmSaidaProduto.Txt(5).Text) = 0 Then
                    FrmSaidaProduto.icms.Text = "18"
              Else
                    FrmSaidaProduto.icms = FrmSaidaProduto.Txt(5).Text
              End If
             End If
         Else
            FrmSaidaProduto.icms = RsProduto!icms
         End If

         
       ' FrmSaidaProduto.txt(5).Text = MostraCliente.TextMatrix(a, 5)
      Case Is = "FrmVales"
        'LcCriSql = "select * from alid009 where codigo=" & MostraCliente.TextMatrix(a - 1, 0)
       ' Set RsProduto = AbreRecordsetLeitura(LcCriSql) ', dbOpenDynaset, dbSeeChanges, dbOptimistic) ', dbOpenDynaset)

        FrmVales.Txt(2).Text = MostraCliente.TextMatrix(a, 1)
        FrmVales.Txt(1).Text = MostraCliente.TextMatrix(a, 0)
        FrmVales.Txt(4).Text = MostraCliente.TextMatrix(a, 8)
       ' RsProduto.FindFirst "cod='" & MostraCliente.TextMatrix(a - 1, 0) & "'"
        FrmVales.valor(0).Text = MostraCliente.TextMatrix(a, 2)
        If Not IsNull(RsProduto!Preco) Then PrecoVendaNormal = CCur(MostraCliente.TextMatrix(a, 2)) / CCur(MostraCliente.TextMatrix(a, 8)) Else PrecoVendaNormal = 0
        ComNormal = CCur(MostraCliente.TextMatrix(a, 8))
        'LcCriterio = "codigo=" & MostraCliente.TextMatrix(a, 0)
        'RsProduto.Find LcCriterio
        If Not RsProduto.EOF Then
           LccriterioUn = "cod='" & RsProduto!UnidMedida & "'"
           RsUnidade.FindFirst LccriterioUn
           If Not RsUnidade.NoMatch Then
              FrmVales.Unidade.Text = RsUnidade!Simbolo
           End If
        End If
        FrmVales.minimo.Text = RsProduto!MinimoVenda & ""
         FrmVales.cst.Text = RsProduto!cst
        If RsProduto!icms = 0 Or IsNull(RsProduto!icms) Then
            If Val(FrmVales.cst.Text) = 60 Or Val(FrmVales.cst.Text) = 160 Or Val(FrmVales.cst.Text) = 260 Then
                    FrmVales.icms.Text = "00"
            Else
              If Len(FrmVales.Txt(5).Text) = 0 Then
                    FrmVales.icms.Text = "18"
              Else
                    FrmVales.icms = FrmVales.Txt(5).Text
              End If
             End If
         Else
            FrmVales.icms = RsProduto!icms
         End If
         
         If Not IsNull(RsProduto!MinimoVenda) Then PrecoMimimodeVendaAlterado = CCur(MostraCliente.TextMatrix(a, 2)) / CCur(MostraCliente.TextMatrix(a, 8)) Else PrecoMimimodeVendaAlterado = 0
       ' FrmSaidaProduto.txt(5).Text = MostraCliente.TextMatrix(a, 5)
  
    Case Is = "FrmReajustaPreco"
        FrmReajustaPreco.bo.Text = MostraCliente.TextMatrix(a, 1)
        FrmReajustaPreco.Codigo.Text = MostraCliente.TextMatrix(a, 0)
    Case Is = "FrmPedido"
        FrmPedido.Txt(2).Text = MostraCliente.TextMatrix(a, 1)
        FrmPedido.Txt(1).Text = MostraCliente.TextMatrix(a, 0)
        FrmPedido.Txt(4).Text = MostraCliente.TextMatrix(a, 8)
       ' RsProduto.FindFirst "cod='" & MostraCliente.TextMatrix(a - 1, 0) & "'"
        'FrmPedido.Valor(0).Text = MostraCliente.TextMatrix(a, 2)
         ComNormal = CCur(MostraCliente.TextMatrix(a, 8))
         FrmPedido.minimo.Text = RsProduto!MinimoVenda & ""
         If Not IsNull(RsProduto!MinimoVenda) Then PrecoMimimodeVendaAlterado = RsProduto!MPVENDA / CCur(MostraCliente.TextMatrix(a, 8)) Else PrecoMimimodeVendaAlterado = 0
       ' FrmSaidaProduto.txt(5).Text = MostraCliente.TextMatrix(a, 5)
        'FrmSaidaProduto.txt(5).Text = MostraC
        
    Case Is = "Orcamento"

        'lcpesqunidade = "cod='" & MostraCliente.TextMatrix(a, 0) & "'"
        'RsProduto.Find lcpesqunidade
        If Not RsProduto.EOF Then
           orcamento.codigoproduto.Text = RsProduto!Codigo
           orcamento.NomeProduto.Text = RsProduto!Nome
           orcamento.ipi.Text = RsProduto!ipi & ""
           lcprocuraunidade = "cod='" & RsProduto!UnidMedida & "'"
           orcamento.Industria.Text = RsProduto!fornecedor & ""
           RsUnidade.FindFirst lcprocuraunidade
           If Not RsUnidade.NoMatch Then
              orcamento.Unidade.Text = RsUnidade!Simbolo
              orcamento.codigounidade.Text = RsUnidade!cod
           End If
           If Len(RsProduto!Preco) > 0 Then
              orcamento.Unitario.Text = RsProduto!Preco
              orcamento.preconormal.Text = RsProduto!Preco
           Else
              orcamento.Unitario.Text = 0
           End If
        End If
        orcamento.Unidade.SetFocus
        RsProduto.Close
        RsUnidade.Close
     Case Is = "FrmProposta"
        FrmProposta.Txt(2).Text = MostraCliente.TextMatrix(a, 1)
        FrmProposta.Txt(1).Text = MostraCliente.TextMatrix(a, 0)
        FrmProposta.Txt(4).Text = MostraCliente.TextMatrix(a, 8)
        'RsProduto.FindFirst "cod='" & MostraCliente.TextMatrix(a - 1, 0) & "'"
        FrmProposta.valor(0).Text = MostraCliente.TextMatrix(a, 2)
        If Not IsNull(RsProduto!Preco) Then PrecoVendaNormal = CCur(MostraCliente.TextMatrix(a, 2)) / CCur(MostraCliente.TextMatrix(a, 8)) Else PrecoVendaNormal = 0
        ComNormal = CCur(MostraCliente.TextMatrix(a, 8))
        'LcCriterio = "cod='" & MostraCliente.TextMatrix(a, 0) & "'"
        'RsProduto.FindFirst LcCriterio
        If Not RsProduto.EOF Then
           LccriterioUn = "cod='" & RsProduto!UnidMedida & "'"
           RsUnidade.FindFirst LccriterioUn
           If Not RsUnidade.NoMatch Then
              FrmProposta.Unidade.Text = RsUnidade!Simbolo
           End If
        End If
        FrmProposta.minimo.Text = RsProduto!MinimoVenda & ""
        FrmProposta.cst.Text = RsProduto!cst
         
        If RsProduto!icms = 0 Or IsNull(RsProduto!icms) Then
            If Val(FrmProposta.cst.Text) = 60 Or Val(FrmProposta.cst.Text) = 160 Or Val(FrmProposta.cst.Text) = 260 Then
                    FrmProposta.icms.Text = "00"
            Else
              If Len(FrmProposta.Txt(5).Text) = 0 Then
                    FrmProposta.icms.Text = "18"
              Else
                    FrmProposta.icms = FrmVales.Txt(5).Text
              End If
             End If
         Else
            FrmProposta.icms = RsProduto!icms
         End If
         If Not IsNull(RsProduto!MinimoVenda) Then PrecoMimimodeVendaAlterado = CCur(MostraCliente.TextMatrix(a, 2)) / CCur(MostraCliente.TextMatrix(a, 8)) Else PrecoMimimodeVendaAlterado = 0
       ' FrmSaidaProduto.txt(5).Text = MostraCliente.TextMatrix(a, 5)

        
    End Select
    Unload Me
  GlFormA.SetFocus
End If
End Sub

Private Sub Parte_Click()
Txt.SetFocus
End Sub

Private Sub Parte_KeyPress(KeyAscii As Integer)
If KeyCode = 27 Then Me.Visible = False
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{C}"
End Sub

Private Sub Txt_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then
   MostraCliente.SetFocus
End If
If KeyCode = 27 Then Me.Visible = False
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{C}"

End Sub


