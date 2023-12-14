VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FrmVendas 
   AutoRedraw      =   -1  'True
   Caption         =   "Vendas"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9570
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   9570
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   14
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   5
      Left            =   5040
      TabIndex        =   7
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   4
      Left            =   3000
      TabIndex        =   6
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton CmdBuscaRes 
      Caption         =   "&Busca de Reserva"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3480
      TabIndex        =   26
      Top             =   6000
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton CmdDevolucao 
      Caption         =   "&Devolução"
      Height          =   495
      Left            =   1920
      TabIndex        =   25
      Top             =   6000
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid MsItens 
      Height          =   2415
      Left            =   120
      TabIndex        =   24
      Top             =   2400
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   4260
      _Version        =   393216
      Cols            =   6
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Fechar"
      Height          =   495
      Left            =   360
      TabIndex        =   23
      Top             =   6000
      Width           =   1455
   End
   Begin VB.TextBox txt 
      Enabled         =   0   'False
      Height          =   285
      Index           =   12
      Left            =   4320
      TabIndex        =   21
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox txt 
      Enabled         =   0   'False
      Height          =   285
      Index           =   11
      Left            =   3000
      TabIndex        =   20
      Text            =   "0"
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox txt 
      Enabled         =   0   'False
      Height          =   285
      Index           =   10
      Left            =   1680
      TabIndex        =   12
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox txt 
      Enabled         =   0   'False
      Height          =   285
      Index           =   9
      Left            =   360
      TabIndex        =   11
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox txt 
      Enabled         =   0   'False
      Height          =   285
      Index           =   8
      Left            =   5760
      TabIndex        =   10
      Top             =   5520
      Width           =   975
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   7
      Left            =   480
      TabIndex        =   9
      Top             =   2040
      Width           =   8295
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   6
      Left            =   480
      TabIndex        =   5
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   3
      Left            =   2760
      TabIndex        =   4
      Top             =   720
      Width           =   6015
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   2
      Left            =   1440
      TabIndex        =   3
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   1
      Left            =   4320
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Quantidade"
      Height          =   255
      Index           =   3
      Left            =   3000
      TabIndex        =   29
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "V. Unitario"
      Height          =   255
      Index           =   2
      Left            =   5040
      TabIndex        =   28
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "V. Total"
      Height          =   255
      Index           =   1
      Left            =   7080
      TabIndex        =   27
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label11 
      Caption         =   "Total Geral"
      Height          =   255
      Left            =   4320
      TabIndex        =   22
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label Label10 
      Caption         =   "Acréscimo"
      Height          =   255
      Left            =   3120
      TabIndex        =   19
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Descontos"
      Height          =   255
      Left            =   1680
      TabIndex        =   18
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Valor Produtos"
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Itens"
      Height          =   255
      Left            =   5880
      TabIndex        =   16
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Cod. Produto"
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   15
      Top             =   1320
      Width           =   930
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   9120
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label3 
      Caption         =   "Responsável"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Data Venda"
      Height          =   255
      Index           =   0
      Left            =   3000
      TabIndex        =   13
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Venda Nº"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "FrmVendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type mtDadosVenda
      codigo As String
      Descricao As String
      Unitario As Currency
      total As Currency
      Quantidade As Currency
      desconto  As Currency
      acrescimo As Currency
      item As Long
      Custo As Currency
End Type
      
Private MtVendas() As mtDadosVenda
Private LcIndicevenda, LcDiaVencimento, LcTeclaNovo As String
Private LcFechaitem, LcItem, LcCodigo, LcDesconto, LcTamanhoGrid As Long
Private LcCusto As Currency
Private LcData As Date
Private RsVendas As Recordset, RsSubvenda As Recordset, RsConvenio As Recordset


Private Sub CmdDevolucao_Click()
On Error Resume Next
FrmDevolucao.Show
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Function GeraGrid()
MsItens.ColWidth(0) = 400
MsItens.ColWidth(1) = 1300
MsItens.ColWidth(2) = 3000
MsItens.ColWidth(3) = 1300
MsItens.ColWidth(4) = 1300
MsItens.ColWidth(4) = 1300
MsItens.TextMatrix(0, 0) = "Item"
MsItens.TextMatrix(0, 1) = "Código"
MsItens.TextMatrix(0, 2) = "Descrição"
MsItens.TextMatrix(0, 3) = "Quantidade"
MsItens.TextMatrix(0, 4) = "V. Unit."
MsItens.TextMatrix(0, 5) = "V. Total"
LcTamanhoGrid = 1
End Function

Function Inicializavenda()
Dim RsFuncao As Recordset

Set Dbbase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
Set RsFuncao = Dbbase.OpenRecordset("TeclasFuncaoLoc", dbOpenDynaset)
LcCriterio = "FuncãodaTecla='Novo Lançamento'"
RsFuncao.FindFirst LcCriterio
If Not RsFuncao.NoMatch Then
   LcTeclaNovo = tecla(RsFuncao!CodigoTecla)
End If
GlDescontos = 0
RsFuncao.Close
Dbbase.Close
AbreBase
GeraGrid
End Function
Function tecla(LCCCodigo As Integer)

Select Case LCCCodigo
Case Is = 113
     tecla = "F2"
Case Is = 114
     tecla = "F3"
Case Is = 115
     tecla = "F4"
Case Is = 116
     tecla = "F5"
Case Is = 117
     tecla = "F6"
Case Is = 118
     tecla = "F7"
Case Is = 119
     tecla = "F8"
Case Is = 120
     tecla = "F9"
Case Is = 121
     tecla = "F10"
Case Is = 122
     tecla = "F11"
Case Is = 123
     tecla = "F12"
End Select
End Function
Function Novo()
On Error Resume Next
Dim a As Integer
For a = 0 To 14
    Txt(a).Text = ""
Next
err = 0
RsVendas.MoveLast
If err <> 0 Then
   LcCodigo = 1
   MsgBox err.Description & err
Else
   LcCodigo = RsVendas!venda + 1
End If
LcData = GlDataSistema
Txt(0).Text = LcCodigo
Txt(1).Text = LcData
LcTamanho = 0
ReDim Preserve MtProduto(LcTamanho)
MsItens.Rows = LcTamanho + 1
'== Cria novo Registro
RsVendas.AddNew
RsVendas!venda = LcCodigo
RsVendas!DATAVENDA = LcData
RsVendas.Update
LcItem = 0
LcTamanhoGrid = 1
Txt(2).SetFocus
End Function
Function BuscaConvenio()

Set Dbbase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
Set RsConvenio = Dbbase.OpenRecordset("Convenio", dbOpenTable, dbSeeChanges, dbOptimistic)
RsConvenio.Index = "Codigo"
GlChave = LcCodigoConvenio
RsConvenio.Seek "=", GlChave
If Not RsConvenio.NoMatch Then
   LcDesconto = RsConvenio!desconto
Else
   LcDesconto = 0
End If
RsConvenio.Close

End Function
Function AbreBase()
On Error Resume Next
Set Dbbase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
Set RsVendas = Dbbase.OpenRecordset("VendaPrincipal", dbOpenTable, dbSeeChanges, dbOptimistic)
Set RsSubvenda = Dbbase.OpenRecordset("vendaSecundario", dbOpenTable, dbSeeChanges, dbOptimistic)
LcIndicevenda = "Locacao"
RsVendas.Index = LcIndicevenda
RsSubvenda.Index = LcIndicevenda
End Function



Private Sub Form_Activate()
Set GlFormA = Me
GlCarregado = True
'txt(15).Text = GlDiasDevolucao
End Sub

Private Sub Form_Load()
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2

Inicializavenda
End Sub
Function PesquisaFuncionario() As Integer
LcIndice = "Codigo"
Call AbreBanco(Funcionario)
GlChave = Txt(2).Text
If AchaReg(1) Then
   Txt(3).Text = RsAtual!Nome
   PesquisaFuncionario = True
Else
   MsgBox "O Código " & GlChave & " Não Foi Cadastrado...", 64, "Aviso"
   PesquisaFuncionario = False
End If
FechaBanco
End Function

Function LimpaControle()
On Error Resume Next
Dim a As Integer
For a = 0 To 13
    Txt(a).Text = ""
Next
LcTotal = 0
GlAcrescimo = 0
LcDescontos = 0
LcTamanho = 0
GlDescontos = 0
ReDim Preserve MtProduto(LcTamanho)
MsItens.Rows = LcTamanho + 1
Txt(2).SetFocus
End Function



Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

FechaBanco
GlCarregado = False
If Len(Trim(GlFormInicial)) = 0 Then alternativo.Show
End Sub

Function VerificaNova()
If Len(Trim(Txt(0).Text)) = 0 Then
   MsgBox "Para Nova Venda Digite " & LcTeclaNovo, 48, "Aviso"
   Screen.ActiveControl.Text = ""
End If
End Function

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   Select Case Index
          Case Is = 2
               Txt(6).SetFocus
          Case Is = 3
               Txt(4).SetFocus
          Case Is = 4
               Txt(5).SetFocus
          Case Is = 5
               Txt(14).SetFocus
          Case Is = 6
               Txt(4).SetFocus
               
          Case Is = 14
               Txt(6).SetFocus
   End Select
Else
 Call PresTecla(KeyCode)
End If
End Sub
Function PresTecla(LcTecla As Integer)
If LcTecla = 112 Then 'Chamou a ajuda
   FrmAjuda.Show , Me
   Exit Function
End If

If LcTecla >= 113 And LcTecla <= 123 Then 'Pressionou Uma Tecla de Função
   Set Dbbase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
   Set RsFuncao = Dbbase.OpenRecordset("TeclasFuncaoLoc", dbOpenDynaset)
   LcCriterio = "CodigoTecla=" & LcTecla
   RsFuncao.FindFirst LcCriterio
  
   If Not RsFuncao.NoMatch Then
      Select Case RsFuncao!FuncãodaTecla
             Case Is = "Pesquisa"
                   MsgBox "Função Não disponivel no Cadastro de Venda.", 48, "Aviso"
             Case Is = "Acréscimo"
                  If LcTamanhoGrid = 1 Then
                     MsgBox "Não Foi Efetuada Nenhuma Venda para dar Acrescimo...", 48, "Aviso"
                     Exit Function
                  End If
                  FrmAcrescimo.Show , Me
             Case Is = "Desconto"
                  If LcTamanhoGrid = 1 Then
                     MsgBox "Não Foi Efetuada Nenhuma Venda para dar Desconto...", 48, "Aviso"
                     Exit Function
                  End If
                  FrmDesconto.Show , Me
             Case Is = "Fecha Locação"
                  If Len(Trim(Txt(12).Text)) = 0 Then
                     MsgBox "Não Foi Efetuada Nenhuma Venda para se Fechar...", 48, "Aviso"
                     Exit Function
                  Else
                     FrmFechaVenda.Show , Me
                  End If
             Case Is = "Cancela Item"
                  If LcTamanhoGrid = 1 Then
                     MsgBox "Não Exite Item Para dar Desconto...", 48, "Aviso"
                     Exit Function
                  End If
                  FrmCancelaItem.Show , Me
             Case Is = "Cancela Locação"
                  CancelaVenda
           
             Case Is = "Pesquisa Item"
                 If Len(Trim(Txt(0).Text)) = 0 Then
                      MsgBox "A Venda Ainda Não foi Inicializada...", 48, "Aviso"
                 Else
                      frmPesquisaProduto.Show , Me
                    
                 End If
             Case Is = "Pesquisa Item"
                  MsgBox "Ainda Não Implementado..."
             
             Case Is = "Novo Lançamento"
                  Novo
             Case Else
                  MsgBox "Esta Tecla Não Foi Programada...", 64, "Aviso"
      End Select
   End If
   
   RsFuncao.Close
Else
  
  Call VerificaNova
End If
End Function
Function Calcula()
Dim LcUnitario, LcTotal As Currency
Dim LcQuantidade As Long
LcUnitario = CCur(Txt(5).Text)
LcQuantidade = CLng(Txt(4).Text)
LcTotal = LcUnitario * LcQuantidade
Txt(14).Text = Format(LcTotal, "Currency")
End Function
Private Sub Txt_LostFocus(Index As Integer)
On Error GoTo ErrLog

Select Case Index
       Case Is = 2
            
            If Screen.ActiveControl.Index = 6 Then
               If Not PesquisaFuncionario Then
                  Txt(Index).SetFocus
               End If
            End If
        Case Is = 4
            Calcula
       Case Is = 5
            
            Calcula
            If LcFechaitem Then
               MontaMatriz
            End If
            Txt(Index).Text = ""
            Txt(14).Text = ""
            Txt(5).Text = ""
            Txt(7).Text = ""
            Txt(6).Text = ""
            Txt(4).Text = ""
            Txt(6).SetFocus
       Case Is = 6
           'Screen.ActiveControl.Index
                     
            
           If Screen.ActiveControl.Index = 4 Then
           
               If Not pesquisaProduto Then
                  Txt(Index).SetFocus
                  LcFechaitem = False
               Else
                 LcFechaitem = True
               End If
            End If
       Case Is = 14
        
           
End Select
Exit Sub
ErrLog:
Exit Sub
End Sub
Function MontaMatriz()
On Error GoTo Erromonta

LcItem = LcItem + 1
ReDim Preserve MtVendas(LcTamanho)
MtVendas(LcTamanho).codigo = Txt(6).Text
MtVendas(LcTamanho).Descricao = Txt(7).Text
MtVendas(LcTamanho).item = Right("00" & LcItem, 2)
MtVendas(LcTamanho).Unitario = CCur(Txt(5).Text)
MtVendas(LcTamanho).Quantidade = CLng(Txt(4).Text)
MtVendas(LcTamanho).desconto = (LcDesconto / 100) * RsAtual!Valorvenda
MtVendas(LcTamanho).total = MtVendas(LcTamanho).Unitario * MtVendas(LcTamanho).Quantidade
MtVendas(LcTamanho).acrescimo = CCur(GlAcrescimo)

MtVendas(LcTamanho).Custo = LcCusto
LcTamanho = LcTamanho + 1
EscreveGrid

Totaliza
Exit Function
Erromonta:

Resume Next

End Function
Function CancelaVenda()
On Error Resume Next
Dim LcResposta As Integer
If LcTamanhoGrid = 1 Then
   MsgBox "Não Foi Efetuada Nenhuma Venda Para Ser Cancelada...", 48, "Aviso"
   Exit Function
End If
LcResposta = MsgBox("Confirma o Cancelamento desta Venda ? ", 36, "Confirmação")
If LcResposta = 6 Then
   
   RsVendas.Delete
   'RsVendas.Close
   LimpaControle
End If

End Function
Function EscreveGrid()
Dim LcTam As Long
On Error GoTo ErroEscreve

LcTam = LcTamanho - 1
LcTamanhoGrid = LcTamanhoGrid + 1
'LcTamanho = LcTamanho + 1
MsItens.Rows = LcTamanhoGrid
MsItens.ColAlignment(0) = 1
MsItens.ColAlignment(1) = 7
MsItens.ColAlignment(2) = 1
MsItens.ColAlignment(3) = 7
MsItens.ColAlignment(4) = 7
MsItens.ColAlignment(5) = 7

MsItens.TextMatrix(LcTamanhoGrid - 1, 0) = MtVendas(LcTam).item
MsItens.TextMatrix(LcTamanhoGrid - 1, 1) = CStr(MtVendas(LcTam).codigo)
MsItens.TextMatrix(LcTamanhoGrid - 1, 2) = MtVendas(LcTam).Descricao
MsItens.TextMatrix(LcTamanhoGrid - 1, 3) = MtVendas(LcTam).Quantidade
MsItens.TextMatrix(LcTamanhoGrid - 1, 4) = Format(CStr(MtVendas(LcTam).Unitario), "Currency")
MsItens.TextMatrix(LcTamanhoGrid - 1, 5) = Format(MtVendas(LcTam).total, "Currency")
Exit Function
ErroEscreve:
MsgBox err.Description
End Function
Function pesquisaProduto()
LcIndice = "Codigo"
Call AbreBanco(produto)
GlChave = Txt(6).Text

If AchaReg(1) Then

   pesquisaProduto = True
   Txt(7).Text = RsAtual!produto
   Txt(5).Text = RsAtual!PrecoVenda
   LcCusto = RsAtual!Custo
Else
   MsgBox "O Código " & GlChave & " Não Foi Cadastrado...", 64, "Aviso"
   pesquisaProduto = False
End If

End Function
Function RemontaIndice()
Dim a, b, LcAchou As Integer

LcItem = 0
b = 0
For a = 0 To LcTamanho - 1
    If Len(Trim(MtVendas(a).codigo)) > 0 Then
       LcItem = LcItem + 1
       MtVendas(a).item = Right("00" & LcItem, 2)
       b = b + 1
       MsItens.Rows = b + 1
       MsItens.ColAlignment(0) = 1
       MsItens.ColAlignment(1) = 7
       MsItens.ColAlignment(2) = 1
       MsItens.ColAlignment(3) = 7
       MsItens.ColAlignment(4) = 7
       MsItens.ColAlignment(5) = 7
       
       MsItens.TextMatrix(b, 0) = MtVendas(a).item
       MsItens.TextMatrix(b, 1) = CStr(MtVendas(a).codigo)
       MsItens.TextMatrix(b, 2) = MtVendas(a).Descricao
       MsItens.TextMatrix(b, 3) = MtVendas(a).Quantidade
       MsItens.TextMatrix(b, 4) = Format(CStr(MtVendas(a).Unitario), "Currency")
       MsItens.TextMatrix(b, 5) = Format(MtVendas(a).total, "Currency")
       LcAchou = True
    End If
Next
If Not LcAchou Then
   MsItens.Rows = 1
End If
Totaliza
LcTamanhoGrid = b + 1
End Function
Function Totaliza()
Dim LcTotal, LcDescontos, LcACrescimo, LcTotalGeral As Currency
Dim LcQuan, LcTotalItem, LcTotalItems, a As Long
Dim LcDataAt As Date
LcTotalItems = 0

If IsEmpty(GlAcrescimo) Then GlAcrescimo = 0
If IsEmpty(LcACrescimo) Then LcACrescimo = 0
For a = 0 To LcTamanho - 1
  If Len(Trim(MtVendas(a).codigo)) > 0 Then
    LcTotalItems = Val(Txt(8).Text) + Val(Txt(4).Text)
    LcTotal = LcTotal + MtVendas(a).total
    LcDescontos = LcDescontos + MtVendas(a).desconto
    LcACrescimo = LcACrescimo + MtVendas(a).acrescimo
  End If
Next
LcDescontos = LcDescontos + GlDescontos
LcACrescimo = LcACrescimo + GlAcrescimo
LcTotalGeral = LcTotal + GlAcrescimo - LcDescontos

Txt(9).Text = Format(LcTotal, "Currency")
Txt(10).Text = Format(LcDescontos, "Currency")
Txt(11).Text = Format(GlAcrescimo, "Currency")
Txt(12).Text = Format(LcTotalGeral, "Currency")
Txt(8).Text = LcTotalItems

End Function
Function ExcluiItem(LcItem As String)
Dim a, LcEncontrado As Integer

For a = 0 To LcTamanho - 1
    If MtVendas(a).item = LcItem Then
       MtVendas(a).codigo = ""
       MtVendas(a).acrescimo = 0
       MtVendas(a).desconto = 0
       MtVendas(a).Descricao = ""
       MtVendas(a).item = 0
       MtVendas(a).Quantidade = 0
       MtVendas(a).total = 0
       MtVendas(a).Unitario = 0
       MtVendas(a).Custo = 0
       LcEncontrado = True
       Exit For
    Else
       LcEncontrado = False
    End If
Next
If Not LcEncontrado Then
   MsgBox "O Item " & LcItem & " Não Foi Encontrado...", 48, "Aviso"
Else
     RemontaIndice
End If
End Function
Function Fechamento(Tipo As String)
On Error Resume Next
Dim a As Integer
Call CriaHistorico(Tipo)
Call GeraVenda(Tipo)
geraSubVenda

GeraCaixa
For a = 0 To 14
    Txt(a).Text = ""
    MsItens.Rows = 1
    LcTamanhoGrid = 1
Next

GlLiberaVenda = True
End Function

Function geraSubVenda()
On Error GoTo errSubVenda
Dim RsSub As Recordset
Dim LcLocacao As Long
Dim a As Integer
LcLocacao = Val(Txt(0).Text)
Set Dbbase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
Set RsSub = Dbbase.OpenRecordset("VendaSecundario", dbOpenTable, dbSeeChanges, dbOptimistic)
For a = 0 To LcTamanho - 1
   RsSub.AddNew
   RsSub!venda = LcLocacao
   RsSub!codigoproduto = MtVendas(a).codigo
   RsSub!DescricaoProd = MtVendas(a).Descricao
   RsSub!Quantidade = MtVendas(a).Quantidade
   RsSub!ValorUnitario = MtVendas(a).Unitario
   RsSub!VALORTOTAL = MtVendas(a).total
   RsSub!Custo = MtVendas(a).Custo
   RsSub!item = MtVendas(a).item
   RsSub!Lucro = (MtVendas(a).Unitario - MtVendas(a).Custo) * MtVendas(a).Quantidade
   RsSub.Update
Next
RsSub.Close
Exit Function

errSubVenda:
If err = 3075 Then
   MsgBox "Valor digitado Incorreto." & Chr(13) & "Caaso Seja data, utilize o formato: " & Chr(13) & "DD/MM/YYYY", 48, "Aviso"
   Exit Function
End If
Select Case ErrosSistema
       Case Is = 0
          Exit Function
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
Function CriaHistorico(Tipo As String)
On Error Resume Next
Dim RsHistorico As Recordset

Dim LcLocacao As Long
Dim a As Integer
LcLocacao = Val(Txt(0).Text)
Set Dbbase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
Set RsHistorico = Dbbase.OpenRecordset("HistoricoVenda", dbOpenTable, dbSeeChanges, dbOptimistic)
For a = 0 To LcTamanho - 1
   RsHistorico.AddNew
   RsHistorico!venda = LcLocacao
   RsHistorico!Responsavel = Txt(2).Text
   RsHistorico!Data = Format(Txt(1).Text, "dd/mm/yyyy")
   RsHistorico!produto = MtVendas(a).codigo
   RsHistorico!DescricaoProduto = MtVendas(a).Descricao
   RsHistorico!Quantidade = MtVendas(a).Quantidade
   RsHistorico!Unitario = MtVendas(a).Unitario
   RsHistorico!total = MtVendas(a).total
   RsHistorico!Custo = MtVendas(a).Custo
   RsHistorico!Lucro = (MtVendas(a).Unitario - MtVendas(a).Custo) * MtVendas(a).Quantidade
   RsHistorico!desconto = Txt(10).Text
   RsHistorico!acrescimo = Txt(11).Text
   RsHistorico!TipoRecebimento = Tipo
   
   RsHistorico.Update
Next
RsHistorico.Close
End Function
Function GeraVenda(Tipo As String)
On Error Resume Next
Dim RsPrincipal As Recordset
Dim LcLocacao As Long
Dim a As Integer
Dim LcData As Date
LcData = Format(Txt(1).Text, "dd/mm/yyyy")
LcLocacao = Val(Txt(0).Text)

Set Dbbase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
Set RsPrincipal = Dbbase.OpenRecordset("VendaPrincipal", dbOpenTable, dbSeeChanges, dbOptimistic)
RsPrincipal.Index = "Locacao"
RsPrincipal.Seek "=", LcLocacao
If Not RsPrincipal.NoMatch Then
   RsPrincipal.Edit
Else
   RsPrincipal.AddNew
End If
   RsPrincipal!venda = LcLocacao
   RsPrincipal!DATAVENDA = LcData
   RsPrincipal!Responsavel = Val(Txt(2).Text)
   RsPrincipal!Descontos = CCur(Txt(10).Text)
   RsPrincipal!TotaVenda = CCur(Txt(12).Text)
   RsPrincipal!QuantProduto = Val(Txt(8).Text)
   RsPrincipal!acrescimo = CCur(Txt(11).Text)
   RsPrincipal!TipoRecebimento = Tipo
   
   RsPrincipal.Update

RsPrincipal.Close

End Function
Function GeraCaixa()
On Error GoTo erroGeraCaixa
Dim RsCaixa As Recordset
Dim LcLocacao As Long
Dim a As Integer
Dim LcData As Date
Dim LcSaldoTotal, LcSaldoDia, LcTotalEntrada, LcDescontos, LcAcrescimos As Currency
Dim LcTotalVendas, LcCaixaAnterior As Currency
Dim LcItem As Long

LcTotalVendas = CCur(Txt(12).Text)
LcDescontos = CCur(Txt(10).Text)
LcAcrescimos = CCur(Txt(11).Text)
LcItem = CLng(Txt(8).Text)
LcData = Format(Txt(1).Text, "dd/mm/yyyy")
'=verifica Brindes

Set Dbbase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
Set RsCaixa = Dbbase.OpenRecordset("Caixa", dbOpenTable, dbSeeChanges, dbOptimistic)

RsCaixa.Index = "DataDoLancamento"
RsCaixa.Seek "=", LcData

If Not RsCaixa.NoMatch Then
   LcSaldoDia = LcTotalVendas + RsCaixa!Saldo
   LcTotalVendas = LcTotalVendas + RsCaixa!EntradaVenda
   LcDescontos = LcDescontos + RsCaixa!Descontos
   LcAcrescimos = LcAcrescimos + RsCaixa!Acrescimos
    
   RsCaixa.Edit
Else
   LcSaldoDia = LcTotalVendas
   LcTotalVendas = LcTotalVendas
   RsCaixa.AddNew
End If
RsCaixa!EntradaVenda = LcTotalVendas
RsCaixa!Saldo = LcSaldoDia
RsCaixa!Descontos = LcDescontos
RsCaixa!Acrescimos = LcAcrescimos
RsCaixa!DataMovimento = LcData
RsCaixa.Update
RsCaixa.Close
Exit Function

erroGeraCaixa:
If err = 3075 Then
   MsgBox "Valor digitado Incorreto." & Chr(13) & "Caaso Seja data, utilize o formato: " & Chr(13) & "DD/MM/YYYY", 48, "Aviso"
   Exit Function
End If
Select Case ErrosSistema
       Case Is = 0
          Exit Function
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

