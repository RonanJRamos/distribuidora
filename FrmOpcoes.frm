VERSION 5.00
Begin VB.Form FrmOpcoes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opções do Sisterma"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8415
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "&Impressoras F3"
      Height          =   495
      Left            =   6720
      TabIndex        =   52
      Top             =   0
      Width           =   1335
   End
   Begin VB.TextBox msg 
      Height          =   375
      Left            =   120
      MaxLength       =   43
      TabIndex        =   46
      Top             =   7560
      Width           =   8055
   End
   Begin VB.Frame Frame8 
      Caption         =   "Comissão"
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   5640
      TabIndex        =   42
      Top             =   2400
      Width           =   2535
      Begin VB.CheckBox comissao 
         Caption         =   "Múltiplas Comissões"
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Orçamento"
      ForeColor       =   &H00C00000&
      Height          =   1815
      Left            =   2880
      TabIndex        =   36
      Top             =   5640
      Width           =   5295
      Begin VB.ComboBox portaorcamento 
         Height          =   315
         Left            =   600
         TabIndex        =   43
         Top             =   360
         Width           =   2175
      End
      Begin VB.CheckBox Cliente 
         Caption         =   "Escolhe Cliente"
         Height          =   255
         Left            =   3600
         TabIndex        =   51
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CheckBox Vendedor 
         Caption         =   "Escolhe Vendedor"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CheckBox Transp 
         Caption         =   "Dados Transp."
         Height          =   255
         Left            =   -480
         TabIndex        =   49
         Top             =   1800
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox margem 
         Height          =   285
         Left            =   3600
         TabIndex        =   33
         Text            =   "0"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox linhasorcamento 
         Height          =   375
         Left            =   2640
         TabIndex        =   35
         Top             =   720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox colunas 
         Caption         =   "Imprime 40 Colunas"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   840
         Width           =   1815
      End
      Begin VB.Line Line8 
         BorderColor     =   &H80000005&
         X1              =   5280
         X2              =   0
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label Label11 
         Caption         =   "Margem"
         Height          =   255
         Left            =   2880
         TabIndex        =   45
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Linhas ao Terminar"
         Height          =   195
         Left            =   3480
         TabIndex        =   39
         Top             =   840
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.Label Label9 
         Caption         =   "Pular"
         Height          =   255
         Left            =   2160
         TabIndex        =   38
         Top             =   840
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Porta"
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Cálculo De Código"
      ForeColor       =   &H00FF0000&
      Height          =   1575
      Left            =   120
      TabIndex        =   31
      Top             =   5640
      Width           =   2655
      Begin VB.CheckBox CodigoFornecedor 
         Caption         =   "Fornecedor"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   960
         Width           =   1575
      End
      Begin VB.CheckBox CodigoCliente 
         Caption         =   "Cliente"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   600
         Width           =   2295
      End
      Begin VB.CheckBox codigoproduto 
         Caption         =   "Produto"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.TextBox Boleto 
      Height          =   315
      Left            =   840
      TabIndex        =   2
      Top             =   1560
      Width           =   3735
   End
   Begin VB.TextBox Nota 
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Top             =   1200
      Width           =   3735
   End
   Begin VB.Frame Frame5 
      Caption         =   "Utilização do Sistema"
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   5640
      TabIndex        =   27
      Top             =   1320
      Width           =   2535
      Begin VB.CheckBox Representante 
         Caption         =   "Representação"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   600
         Width           =   1935
      End
      Begin VB.CheckBox Comercio 
         Caption         =   "Comercial"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Preços"
      ForeColor       =   &H00FF0000&
      Height          =   2055
      Left            =   120
      TabIndex        =   23
      Top             =   3480
      Width           =   2655
      Begin VB.CheckBox Check3 
         Caption         =   "Altera  Minimo de Venda Na Alteração de Preço "
         Height          =   615
         Left            =   120
         TabIndex        =   26
         Top             =   1320
         Width           =   2415
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Alterar Lucro Na Alteração de Preço"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   840
         Width           =   2415
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Altera na Digitação Lucro no cadastro de Produto."
         Height          =   615
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Lançamentos Contas a Receber"
      ForeColor       =   &H00FF0000&
      Height          =   2415
      Left            =   5640
      TabIndex        =   18
      Top             =   3120
      Width           =   2535
      Begin VB.CheckBox CaixaSaida 
         Caption         =   "Atualizar Caixa Por Venda a Vista."
         Height          =   675
         Left            =   120
         TabIndex        =   21
         Top             =   1560
         Width           =   2175
      End
      Begin VB.CheckBox VistaSaida 
         Caption         =   "Atualizar Por Venda a Vista."
         Height          =   615
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   2055
      End
      Begin VB.CheckBox FaturaSaida 
         Caption         =   "Atualizar  Por Venda  Faturada."
         Height          =   615
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Lançamentos Contas a Pagar"
      ForeColor       =   &H00FF0000&
      Height          =   2895
      Left            =   2880
      TabIndex        =   14
      Top             =   2640
      Width           =   2655
      Begin VB.CheckBox CaixaEntrada 
         Caption         =   "Atualizar Caixa Por Nota de Entrada a Vista."
         Height          =   675
         Left            =   240
         TabIndex        =   17
         Top             =   1560
         Width           =   1695
      End
      Begin VB.CheckBox VistaEntrada 
         Caption         =   "Atualizar Por Notas de Entradas a Vista."
         Height          =   615
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Width           =   2055
      End
      Begin VB.CheckBox FaturaEntrada 
         Caption         =   "Atualizar  Por Nota de Entrada Faturada."
         Height          =   615
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Confirmações"
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Width           =   2655
      Begin VB.CheckBox Excluir 
         Caption         =   " Antes de Excluir"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   2175
      End
      Begin VB.CheckBox Alterar 
         Caption         =   " Antes de Alterar"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   2295
      End
      Begin VB.CheckBox Incluir 
         Caption         =   "Confirma Antes de Incluir"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   840
         Visible         =   0   'False
         Width           =   2415
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar F10"
      Height          =   495
      Left            =   5280
      TabIndex        =   6
      Top             =   600
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Salvar F2"
      Height          =   495
      Left            =   5280
      TabIndex        =   5
      Top             =   0
      Width           =   1335
   End
   Begin VB.TextBox txt 
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   450
      Width           =   1095
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mensagem No Pedido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   120
      TabIndex        =   48
      Top             =   7320
      Width           =   2040
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Linhas ao Terminar"
      Height          =   195
      Left            =   6600
      TabIndex        =   47
      Top             =   6600
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Para Impressora de Rede Utilize o caminho completo.   Ex: \\Serv\Impressora"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Left            =   120
      TabIndex        =   30
      Top             =   2040
      Width           =   4620
   End
   Begin VB.Line Line7 
      BorderColor     =   &H8000000E&
      X1              =   8400
      X2              =   5040
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line6 
      BorderColor     =   &H8000000E&
      X1              =   5040
      X2              =   5040
      Y1              =   0
      Y2              =   1200
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Impressão Nota Saída  e Boleto "
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   0
      Width           =   2535
   End
   Begin VB.Line Line5 
      X1              =   4680
      X2              =   2640
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000005&
      X1              =   4680
      X2              =   4680
      Y1              =   1920
      Y2              =   120
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   4680
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   1920
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Portas Impressora"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Width           =   1260
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      X1              =   120
      X2              =   4680
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label4 
      Caption         =   "Boleto"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Nota"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "No inicio da Nota Fiscal"
      Height          =   195
      Left            =   2040
      TabIndex        =   4
      Top             =   540
      Width           =   1680
   End
   Begin VB.Label Label1 
      Caption         =   "Pular"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   615
   End
End
Attribute VB_Name = "FrmOpcoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LcAchou As Integer, LcAtivos As Integer
Dim a As Integer
Private Sub Alterar_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 113 Then SendKeys "%+{S}"
If KeyCode = 114 Then SendKeys "%+{I}"
End Sub

Private Sub boleto_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 113 Then SendKeys "%+{S}"
If KeyCode = 114 Then SendKeys "%+{I}"
End Sub

Private Sub boleto_LostFocus()
If Boleto.Text = Nota.Text Then
   MsgBox "A porta " & Boleto.Text & " Já está definida para a Nota Fiscal.", 64, "Aviso"
   Boleto.SetFocus
End If

End Sub

Private Sub CaixaEntrada_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   SendKeys "{TAB}"
Else
    Call Teclas(KeyCode)
End If
End Sub

Private Sub CaixaSaida_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   SendKeys "{TAB}"
Else
    Call Teclas(KeyCode)
End If
End Sub

Private Sub Check1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 113 Then SendKeys "%+{S}"
If KeyCode = 114 Then SendKeys "%+{I}"

End Sub

Private Sub Check2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   SendKeys "{TAB}"
Else
    Call Teclas(KeyCode)
End If

End Sub

Private Sub Check3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   SendKeys "{TAB}"
Else
    Call Teclas(KeyCode)
End If

End Sub

Private Sub colunas_Click()
If colunas = 1 Then
   Label9.Visible = True
   Label10.Visible = True
   linhasorcamento.Visible = True
Else
   Label9.Visible = False
   Label10.Visible = False
   linhasorcamento.Visible = False
End If
End Sub

Private Sub colunas_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 113 Then SendKeys "%+{S}"
End Sub

Private Sub Command1_Click()
'On Error Resume Next
Dim RsOpcoes As Recordset
Dim LcCap As String
LcCap = Me.Caption
NumeroDoArquivo = FreeFile
NumeroMsg = NumeroDoArquivo + 1
Me.Caption = "Aguarde,Atualizando Informações..."

For a = Len(GLBase) To 1 Step -1
    letra = Mid(GLBase, a, 1)
    If letra = "\" Then
       LcArqMsg = Mid(GLBase, 1, a) & "msg.txt"
       Exit For
    End If
Next
Open LcArqMsg For Output As #NumeroMsg
Print #NumeroMsg, msg.Text
Close #NumeroMsg

Open App.Path & "\opcao.txt" For Output As #NumeroDoArquivo
If Len(Txt.Text) = 0 Then Txt.Text = "0"
Print #NumeroDoArquivo, Txt.Text
Print #NumeroDoArquivo, Nota.Text
Print #NumeroDoArquivo, Boleto.Text
Print #NumeroDoArquivo, Incluir
Print #NumeroDoArquivo, Alterar
Print #NumeroDoArquivo, Excluir
Print #NumeroDoArquivo, FaturaSaida
Print #NumeroDoArquivo, VistaSaida
Print #NumeroDoArquivo, CaixaSaida
Print #NumeroDoArquivo, FaturaEntrada
Print #NumeroDoArquivo, VistaEntrada
Print #NumeroDoArquivo, CaixaEntrada
Print #NumeroDoArquivo, Check1
Print #NumeroDoArquivo, Check2
Print #NumeroDoArquivo, Check3
Print #NumeroDoArquivo, Comercio
Print #NumeroDoArquivo, Representante
Print #NumeroDoArquivo, codigoproduto
Print #NumeroDoArquivo, portaorcamento
Print #NumeroDoArquivo, colunas
Print #NumeroDoArquivo, linhasorcamento
Print #NumeroDoArquivo, CodigoCliente
Print #NumeroDoArquivo, CodigoFornecedor
Print #NumeroDoArquivo, comissao
Print #NumeroDoArquivo, margem.Text
Print #NumeroDoArquivo, Transp
Print #NumeroDoArquivo, Vendedor
Print #NumeroDoArquivo, cliente

If Len(Txt.Text) > 0 Then GLSaltoLinhaNota = CInt(Txt.Text) Else GLSaltoLinhaNota = 0
GlMsg = msg.Text
GlPortaNota = Nota.Text
GlPortaBoleto = Boleto.Text
GlPortaOrcamento = portaorcamento.Text
If Len(margem.Text) > 0 Then GlMargem = CLng(margem.Text) Else GlMargem = 0
If comissao = 1 Then GlVariasComissao = True Else GlVariasComissao = False
If Incluir = 1 Then GLConfirmaNovo = True Else GLConfirmaNovo = False
If Alterar = 1 Then GlConfirmaAlteracao = True Else GlConfirmaAlteracao = False
If Excluir = 1 Then GlConfirmaExclusao = True Else GlConfirmaExclusao = False
If FaturaSaida = 1 Then GlFaturaSaida = True Else GlFaturaSaida = False
If VistaSaida = 1 Then GlVistaSaida = True Else GlVistaSaida = False
If CaixaSaida = 1 Then GlCaixaSaida = True Else GlCaixaSaida = False
If FaturaEntrada = 1 Then GlFaturaEntrada = True Else GlFaturaEntrada = False
If VistaEntrada = 1 Then GlVistaEntrada = True Else GlVistaEntrada = False
If CaixaEntrada = 1 Then GlCaixaEntrada = True Else GlCaixaEntrada = False
If Check1 = 1 Then GlLucroCad = True Else GlLucroCad = False
If Check2 = 1 Then GlLucroAlteracao = True Else GlLucroAlteracao = False
If Check3 = 1 Then GlMinimoAlteracao = True Else GlMinimoAlteracao = False
If Comercio = 1 Then GlComercio = True Else GlComercio = False
If Representante = 1 Then GlRepresentante = True Else GlRepresentante = False
If codigoproduto = 1 Then GLCalculacodigoProduto = True Else GLCalculacodigoProduto = False
If CodigoCliente = 1 Then GLCalculacodigoCliente = True Else GLCalculacodigoCliente = False
If CodigoFornecedor = 1 Then GLCalculacodigoFornecedor = True Else GLCalculacodigoFornecedor = False


If Transp = 1 Then GlDadosTransportadora = True Else GlDadosTransportadora = False
If Vendedor = 1 Then GlEsclheVendedor = True Else GlEsclheVendedor = False
If cliente = 1 Then GlEscolheCliente = True Else GlEscolheCliente = False

If colunas = 1 Then Gl40colunas = True Else Gl40colunas = False
If Len(linhasorcamento.Text) > 0 Then GlSaltoLinhasOrcamento = CInt(linhasorcamento.Text) Else GlSaltoLinhasOrcamento = 0

Close #NumeroDoArquivo

Me.Caption = LcCap
Unload Me
End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Teclas(KeyCode)
End Sub

Private Sub Command2_Click()
On Error Resume Next
Unload Me

End Sub

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
Call Teclas(KeyCode)
End Sub

Private Sub Command3_Click()
On Error Resume Next
FrmImpressoras.Show , Me
End Sub

Private Sub Excluir_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 113 Then SendKeys "%+{S}"
If KeyCode = 114 Then SendKeys "%+{I}"
End Sub

Private Sub FaturaEntrada_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   SendKeys "{TAB}"
Else
    Call Teclas(KeyCode)
End If
End Sub

Private Sub FaturaSaida_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   SendKeys "{TAB}"
Else
   Call Teclas(KeyCode)
End If
End Sub

Private Sub Form_Activate()
If LcAtivos Then Exit Sub
LcAtivos = True
Set GlFormA = Me
Me.Refresh
Reconfigura
If Not LcAchou Then
   LcResposta = MsgBox("Não Exite Impressoras Cadastradas no Sistema." & Chr(13) & Chr(13) & "Cadastra Agora ?", vbExclamation + vbYesNo, "Aviso")
   If LcResposta = 6 Then
      FrmImpressoras.Show , Me
   End If
End If
End Sub
Function CarregaCombo()
'On Error Resume Next
Dim RsImpressoras As Recordset

AbreBase
Set RsImpressoras = Dbbase.OpenRecordset("Impressoras", dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcAchou = 0
Do Until RsImpressoras.EOF
   LcAchou = -1
   If err > 0 Then Exit Do
   portaorcamento.AddItem RsImpressoras!impressora
   RsImpressoras.MoveNext
Loop

RsImpressoras.Close
Set RsImpressoras = Nothing


End Function
Function Reconfigura()
Txt.Text = GLSaltoLinhaNota
portaorcamento.Text = GlPortaOrcamento
If Len(GlPortaNota) > 0 Then Nota.Text = GlPortaNota
If Len(GlPortaBoleto) > 0 Then Boleto.Text = GlPortaBoleto
If GLConfirmaNovo Then Incluir = 1 Else Incluir = 0
If GlConfirmaAlteracao Then Alterar = 1 Else Alterar = 0
If GlConfirmaExclusao Then Excluir = 1 Else Excluir = 0
If GlFaturaSaida Then FaturaSaida = 1 Else FaturaSaida = 0
If GlVistaSaida Then VistaSaida = 1 Else VistaSaida = 0
If GlCaixaSaida Then CaixaSaida = 1 Else CaixaSaida = 0
If GlFaturaEntrada Then FaturaEntrada = 1 Else FaturaEntrada = 0
If GlVistaEntrada Then VistaEntrada = 1 Else VistaEntrada = 0
If GlCaixaEntrada Then CaixaEntrada = 1 Else CaixaEntrada = 0
If GlLucroCad Then Check1 = 1 Else Check1 = 0
If GlLucroAlteracao Then Check2 = 1 Else Check2 = 0
If GlMinimoAlteracao Then Check3 = 1 Else Check3 = 0
If GlComercio Then Comercio = 1 Else Comercio = 0
If GlRepresentante Then Representante = 1 Else Representante = 0
If GLCalculacodigoProduto Then codigoproduto = 1 Else codigoproduto = 0
If Gl40colunas Then colunas = 1 Else colunas = 0
If GLCalculacodigoCliente Then CodigoCliente = 1 Else CodigoCliente = 0
If GLCalculacodigoFornecedor Then CodigoFornecedor = 1 Else CodigoFornecedor = 0
If GlVariasComissao Then comissao = 1 Else comissao = 0

If GlDadosTransportadora Then Transp = 1 Else Transp = 0
If GlEsclheVendedor Then Vendedor = 1 Else Vendedor = 0
If GlEscolheCliente Then cliente = 1 Else cliente = 0

If Len(GlMargem) > o Then
   margem.Text = GlMargem
Else
   GlMargem = 0
End If
msg.Text = GlMsg
linhasorcamento.Text = GlSaltoLinhasOrcamento
 
End Function

Private Sub Form_Load()
Me.Refresh
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
CarregaCombo
Reconfigura

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
FrmPrincipal.SetFocus
LcAtivos = False
End Sub

Private Sub linhasorcamento_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 113 Then SendKeys "%+{S}"
End Sub

Private Sub margem_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 113 Then SendKeys "%+{S}"
End Sub

Private Sub msg_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 113 Then SendKeys "%+{S}"

End Sub

Private Sub Nota_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 113 Then SendKeys "%+{S}"
If KeyCode = 114 Then SendKeys "%+{I}"
End Sub

Private Sub portaorcamento_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 113 Then SendKeys "%+{S}"
If KeyCode = 114 Then SendKeys "%+{I}"
End Sub

Private Sub Txt_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 113 Then SendKeys "%+{S}"
If KeyCode = 114 Then SendKeys "%+{I}"
End Sub

Private Sub VistaEntrada_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   SendKeys "{TAB}"
Else
   Call Teclas(KeyCode)
End If
End Sub

Private Sub VistaSaida_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   SendKeys "{TAB}"
Else
   Call Teclas(KeyCode)
 End If
End Sub
