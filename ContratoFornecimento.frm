VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form ContratoFornecimento 
   BackColor       =   &H00CDCDAF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contrato de Fornecimento"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9480
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Lancado 
      Height          =   255
      Left            =   4320
      TabIndex        =   28
      Top             =   720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton CmdFechar 
      BackColor       =   &H00B3B386&
      Caption         =   "&Fechar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7080
      Width           =   1935
   End
   Begin VB.CommandButton CmdPesquisar 
      BackColor       =   &H00B3B386&
      Caption         =   "&Pesquisar Contratos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7080
      Width           =   1935
   End
   Begin VB.CommandButton CmdNovo 
      BackColor       =   &H00B3B386&
      Caption         =   "Novo Contrato"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7080
      Width           =   1815
   End
   Begin VB.CommandButton ExcluiItem 
      BackColor       =   &H00B3B386&
      Caption         =   "Excluir Item"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7080
      Width           =   1815
   End
   Begin VB.CommandButton ExcluiContrato 
      BackColor       =   &H00B3B386&
      Caption         =   "&Excluir Contrato"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7080
      Width           =   1815
   End
   Begin VB.CommandButton CmdLancar 
      BackColor       =   &H00B3B386&
      Caption         =   "Lançar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2280
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid Item 
      Height          =   3975
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   7011
      _Version        =   393216
      Rows            =   1
      Cols            =   5
      BackColor       =   -2147483624
      BackColorFixed  =   12369044
      BackColorBkg    =   15198168
   End
   Begin VB.TextBox ValorUnit 
      Height          =   375
      Left            =   7080
      TabIndex        =   7
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox Produto 
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   2280
      Width           =   5295
   End
   Begin VB.TextBox CodProduto 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CDCDAF&
      Caption         =   "Validade"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   680
      Left            =   7200
      TabIndex        =   23
      Top             =   1200
      Width           =   2175
      Begin MSMask.MaskEdBox DataI 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   8
         Mask            =   "99/99/99"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DataF 
         Height          =   375
         Left            =   1080
         TabIndex        =   4
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   8
         Mask            =   "99/99/99"
         PromptChar      =   " "
      End
   End
   Begin VB.TextBox Cliente 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   5535
   End
   Begin VB.TextBox CodCliente 
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Contrato 
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin MSMask.MaskEdBox Data 
      Height          =   375
      Left            =   8280
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin VB.TextBox Codigo 
      Appearance      =   0  'Flat
      BackColor       =   &H00CDCDAF&
      Height          =   375
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Use o Duplo Click para Alterar Valor Unitário"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2520
      TabIndex        =   27
      Top             =   6840
      Width           =   3975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Precione F5 para Selecionar Produto"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2880
      TabIndex        =   26
      Top             =   2040
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Precione F5 para Selecionar Cliente"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1920
      TabIndex        =   25
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
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
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   24
      Top             =   2040
      Width           =   675
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   9480
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
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
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   22
      Top             =   1200
      Width           =   600
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Unit."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   7080
      TabIndex        =   21
      Top             =   2040
      Width           =   915
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   9600
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº Contrato"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   5760
      TabIndex        =   20
      Top             =   1200
      Width           =   1005
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Emissão"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   7320
      TabIndex        =   19
      Top             =   720
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   720
      Width           =   600
   End
   Begin VB.Label Label1 
      BackColor       =   &H00EAEADD&
      Caption         =   " CONTRATO DE FORNECIMENTO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   10335
   End
End
Attribute VB_Name = "ContratoFornecimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Rs      As ADODB.Recordset
Private a       As Integer
Private LcSql   As String

Private Sub Cliente_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 116 Then
    GlEscolhe = 1
    If Len(Cliente.Text) > 0 Then
        GlCriterioSql = "select * From alid001 where RAZAOSOC like '" & UCase(Cliente.Text) & "*'  order by RAZAOSOC"
    Else
        GlCriterioSql = ""
    End If
    FrmBuscaCliente.Show , Me
End If
End Sub

Private Sub CmdFechar_Click()
Unload Me
End Sub

Private Sub CmdSalar_Click()

End Sub

Private Sub CmdLancar_Click()
On Error Resume Next
Dim RSLanca As ADODB.Recordset
Dim LcCod As Integer
If Len(CodCliente.Text) = 0 Then
    MsgBox "Escolha o Cliente para o Contrato.", 64, "Aviso"
    CodCliente.SetFocus
    Exit Sub
End If
If Len(Contrato.Text) = 0 Then
    MsgBox "Entre com o Número do Contrato.", 64, "Aviso"
    Contrato.SetFocus
    Exit Sub
End If
If Not IsDate(DataI.Text) Or Not IsDate(DataF.Text) Then
    MsgBox "Período de Validade do Contrato Inválido.", 64, "Aviso"
    DataI.SetFocus
    Exit Sub
End If
If Len(CodProduto.Text) = 0 Then
    MsgBox "Entre com o Produto de Lançamento.", 64, "Aviso"
    CodProduto.SetFocus
    Exit Sub
End If
If Lancado.Value = 0 Then
    SalvaCabecalho (1)
Else
    SalvaCabecalho (2)
End If
LcSql = "Insert Into ContratoDados(CodContrato,CodProduto,Produto,Valor)Values("
LcSql = LcSql & Codigo.Text & ","
LcSql = LcSql & "'" & CodProduto.Text & "',"
LcSql = LcSql & "'" & Produto.Text & "',"
LcSql = LcSql & "" & Replace(CDbl(ValorUnit.Text), ",", ".") & ")"
afetados = ExecutaSql(LcSql)
LcSql = "Select Codigo FROM ContratoDados where CodContrato=" & Codigo.Text & " order by codigo"
Set RSLanca = AbreRecordset(LcSql)
If Not RSLanca.EOF Then
    RSLanca.MoveLast
    LcCod = RSLanca!Codigo
End If
RSLanca.Close
Set RSLanca = Nothing
a = Item.Rows
Item.Rows = a + 1
Item.TextMatrix(a, 0) = a
Item.TextMatrix(a, 1) = CodProduto.Text
Item.TextMatrix(a, 2) = Produto.Text & ""
Item.TextMatrix(a, 3) = AcertaNumero(CDbl(ValorUnit.Text), 2)
Item.TextMatrix(a, 4) = LcCod
LimpaProduto
End Sub
Sub LimpaProduto()
On Error Resume Next
CodProduto.Text = ""
Produto.Text = ""
ValorUnit.Text = ""
CodProduto.SetFocus
End Sub
Sub SalvaCabecalho(LcAcao As Integer)
If LcAcao = 1 Then
    LcSql = "Insert Into ContratoFornecimento(Data,Contrato,DataI,DataF,CodCliente,Cliente)Values("
    LcSql = LcSql & "#" & Format(Data.Text, "mm/dd/yy") & "#,"
    LcSql = LcSql & "'" & Contrato.Text & "',"
    LcSql = LcSql & "#" & Format(DataI.Text, "mm/dd/yy") & "#,"
    LcSql = LcSql & "#" & Format(DataF.Text, "mm/dd/yy") & "#,"
    LcSql = LcSql & "'" & CodCliente.Text & "',"
    LcSql = LcSql & "'" & Cliente.Text & "')"
Else
    LcSql = "Update ContratoFornecimento set "
    LcSql = LcSql & "Data=#" & Format(Data.Text, "mm/dd/yy") & "#,"
    LcSql = LcSql & "Contrato='" & Contrato.Text & "',"
    LcSql = LcSql & "DataI=#" & Format(DataI.Text, "mm/dd/yy") & "#,"
    LcSql = LcSql & "DataF=#" & Format(DataF.Text, "mm/dd/yy") & "#,"
    LcSql = LcSql & "CodCliente='" & CodCliente.Text & "',"
    LcSql = LcSql & "Cliente='" & Cliente.Text & "' where codigo=" & Codigo.Text & ")"
End If
afetados = ExecutaSql(LcSql)
If Lancado.Value = 0 Then
    LcSql = "Select Codigo FROM ContratoFornecimento order by codigo"
    Set Rs = AbreRecordset(LcSql, True)
    If Not Rs.EOF Then
        Rs.MoveLast
        Codigo.Text = Rs!Codigo
        Lancado.Value = 1
    End If
    Rs.Close
    Set Rs = Nothing
End If
End Sub

Private Sub CmdNovo_Click()
Call LimpaForm
Cliente.SetFocus
End Sub

Private Sub CmdPesquisar_Click()
On Error Resume Next
LcResposta = InputBox("ENTRE COM O NÚMERO DO CONTRATO A PESQUISAR.", "Pesquisa de Contratos")
If Len(LcResposta) = 0 Then Exit Sub
LimpaForm
PesquisaContrato (LcResposta)
End Sub

Private Sub CodCliente_LostFocus()
On Error Resume Next
Dim bb As Database, rsCliente As Recordset
If Len(Trim(CodCliente.Text)) = 0 Then Exit Sub
'===> Verifica se o Valor digitado é Númerico
If GLCalculaCodCliente Then
   If Not IsNumeric(CodCliente.Text) And Len(CodCliente.Text) > 0 Then
      MsgBox "O Código do Cliente deve ser Numérico...", vbExclamation, "Aviso"
      CodCliente.SetFocus
      Exit Sub
   End If
   CodCliente.Text = Right("00000" & CodCliente.Text, 5)
End If

Set bb = OpenDatabase(GLBase, False, False)
Set rsCliente = bb.OpenRecordset("select * From alid001 where CODIGO='" & CodCliente.Text & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)  ', dbOpenDynaset)
If Not rsCliente.EOF Then
   Cliente.Text = rsCliente!razaosoc
    CodProduto.SetFocus
Else
    CodCliente.Text = ""
End If
rsCliente.Close
bb.Close
End Sub
Sub LimpaForm()
On Error Resume Next
LimpaProduto
Codigo.Text = ""
CodCliente.Text = ""
Cliente.Text = ""
Item.Rows = 1
DataI.Text = "  /  /  "
DataF.Text = "  /  /  "
Data.Text = Format(Date, "dd/mm/yy")
Contrato.Text = ""
Lancado.Value = 0
CodCliente.SetFocus
End Sub
Function BuscaProduto()
On Error Resume Next
If Len(Produto.Text) > 0 And Len(CodProduto.Text) = 0 Then Exit Function
Dim LcAchou As Integer
Dim RsProduto As Recordset, RsUnidade As Recordset
AbreBase
Set RsProduto = Dbbase.OpenRecordset("select * From alid009 where cod='" & CodProduto.Text & "'") ', dbOpenDynaset)
If Not RsProduto.EOF Then
   'LcCriterio = "Cod='" & RsProduto!Unimed & "'"
   Set RsUnidade = Dbbase.OpenRecordset("select * From alid004 where cod='" & RsProduto!Unimed & "'")  ', dbOpenDynaset)

   'RsUnidade.FindFirst LcCriterio
   If Not RsUnidade.EOF Then
      LcUnidade = RsUnidade!Simbolo
   End If
   CodProduto.Text = RsProduto!cod
   Produto.Text = RsProduto!Nome
   LcAchou = True
Else
   CodProduto.Text = ""
   Produto.Text = ""
   LcAchou = False
   CodProduto.SetFocus
   MsgBox "Código Não Encontrado...", 64, "Aviso"
End If
If LcAchou Then SendKeys "{TAB}"
RsProduto.Close
RsUnidade.Close
Set RsProduto = Nothing
Set RsUnidade = Nothing
End Function

Private Sub CodProduto_LostFocus()
On Error Resume Next
Dim Rs As ADODB.Recordset
If Len(CodProduto.Text) = 0 Then Exit Sub
Set Rs = AbreRecordset("Select * from produtos where codigo=" & CodProduto.Text, True)
If Not Rs.EOF Then
   Produto.Text = Rs!Nome
   ValorUnit.SetFocus
Else
   Nome.Text = ""
   CodProduto.Text = ""
   CodProduto.SetFocus
End If
Rs.Close
Set Rs = Nothing
End Sub

Private Sub ExcluiContrato_Click()
On Error Resume Next
If Len(Codigo.Text) = 0 Then Exit Sub
LcSql = MsgBox("CONFIRMAR A EXCLUSÃO DESTE CONTRATO?", vbYesNo, "Exclusão de Contrato")
If LcSql = 7 Then Exit Sub
LcSql = "Delete FROM ContratoFornecimento Where Codigo=" & Codigo.Text
afetados = ExecutaSql(LcSql)
LcSql = "Delete FROM ContratoDados Where CodContrato=" & Codigo.Text
afetados = ExecutaSql(LcSql)
LimpaForm
MsgBox "Contrato Excluído com Sucesso.", 64, "Aviso"
End Sub

Private Sub ExcluiItem_Click()
On Error Resume Next
Dim LcAchou As Boolean
LcAchou = False
If Item.Rows = 1 Then Exit Sub
LcResposta = InputBox("ENTRE COM O ITEM A EXCLUIR.", "Exclusão de Produto do Contrato")
If IsNumeric(LcResposta) Then
    For a = 1 To Item.Rows - 1
        If Int(LcResposta) = Int(Item.TextMatrix(a, 0)) Then
            LcAchou = True
            LcSql = "Delete FROM ContratoDados Where Codigo=" & Item.TextMatrix(a, 4)
            afetados = ExecutaSql(LcSql)
            Exit For
        End If
    Next
    If LcAchou Then
        PesquisaContrato (Codigo.Text)
    Else
        MsgBox "Item não encontrado no contrato.", 64, "Aviso"
    End If
End If
End Sub

Private Sub Form_Activate()
Set GlFormA = Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then
    SendKeys "{tab}"
    SendKeys "{HOME}+{END}"
End If
End Sub

Private Sub Form_Load()
GeraGrid
Data.Text = Format(Date, "dd/mm/yy")
End Sub
Sub PesquisaContrato(LcCodigo As String)
On Error Resume Next
Dim LcCarregou As Boolean
LcCarregou = False
If Len(LcCodigo) = 0 Then Exit Sub
Item.Rows = 1
LcSql = "SELECT ContratoFornecimento.Codigo, ContratoFornecimento.Data, ContratoFornecimento.CodCliente, ContratoFornecimento.Cliente, ContratoFornecimento.Contrato,ContratoFornecimento.DataI,ContratoFornecimento.DataF"
LcSql = LcSql & " FROM ContratoFornecimento"
LcSql = LcSql & " Where ContratoFornecimento.Contrato='" & LcCodigo & "'"
Set Rs = AbreRecordset(LcSql, True)
If Not Rs.EOF Then
    Codigo.Text = Rs!Codigo
    Data.Text = Format(Rs!Data, "dd/mm/yy")
    DataI.Text = Format(Rs!DataI, "dd/mm/yy")
    DataF.Text = Format(Rs!DataF, "dd/mm/yy")
    CodCliente.Text = Rs!CodCliente & ""
    Cliente.Text = Rs!Cliente & ""
    Contrato.Text = Rs!Contrato & ""
    LcCarregou = True
End If
If LcCarregou Then
    LcSql = "Select Codigo,CodContrato,CodProduto,Produto,Valor"
    LcSql = LcSql & " FROM contratodados where CodContrato=" & Codigo.Text & " order by codigo"
    Set Rs = AbreRecordset(LcSql, True)
    Do Until Rs.EOF
        a = Item.Rows
        Item.Rows = a + 1
        Item.TextMatrix(a, 0) = a
        Item.TextMatrix(a, 1) = Rs!CodProduto & ""
        Item.TextMatrix(a, 2) = Rs!Produto & ""
        Item.TextMatrix(a, 3) = AcertaNumero(CDbl(Rs!valor), 2)
        Item.TextMatrix(a, 4) = Rs!Codigo & ""
        Rs.MoveNext
    Loop
    Rs.Close
    Set Rs = Nothing
End If
If Not LcCarregou Then
    MsgBox "Contrato não encontrado.", 64, "Aviso"
End If
End Sub

Private Sub Item_DblClick()
On Error Resume Next
a = Item.Col
If a = 3 Then
    a = Item.Row
    LcValor = InputBox("ENTRE COM O NOVO VALOR", "Alteração do Valor Unitário", Item.TextMatrix(a, 3))
    If IsNumeric(LcValor) Then
        Item.TextMatrix(a, 3) = AcertaNumero(CDbl(LcValor), 2)
        LcSql = "Update ContratoDados set Valor=" & Replace(CDbl(LcValor), ",", ".") & " where codigo=" & Item.TextMatrix(a, 4)
        afetados = ExecutaSql(LcSql)
    End If
End If
End Sub

Private Sub produto_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 116 Then
    GlEscolhe = 2
    If Len(Trim(Produto.Text)) > 0 Then
        FrmPesquisaProdutos.txt.Text = Produto.Text
        GlCriterioSql = "select * From Produtos where nome like '" & UCase(Produto.Text) & "%'  order by nome"
    Else
        GlCriterioSql = ""
    End If
    FrmPesquisaProdutos.Show , Me
End If
End Sub

Private Sub produto_LostFocus()
On Error Resume Next
'BuscaProduto
End Sub
Sub GeraGrid()
On Error Resume Next
Item.ColWidth(0) = 500
Item.ColWidth(1) = 800
Item.ColWidth(2) = 6000
Item.ColWidth(3) = 1500
Item.ColWidth(4) = 0
Item.TextMatrix(0, 0) = "Item"
Item.TextMatrix(0, 1) = "Código"
Item.TextMatrix(0, 2) = "Produto"
Item.TextMatrix(0, 3) = "Valor"
Item.TextMatrix(0, 4) = "Codigo"
End Sub
