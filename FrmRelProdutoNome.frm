VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form FrmRelProdutoNome 
   BackColor       =   &H00EAE8DD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Produto por Nome"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6960
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   6960
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00EAE8DD&
      Caption         =   "Desativados"
      Height          =   1335
      Left            =   4320
      TabIndex        =   19
      Top             =   1800
      Width           =   2535
      Begin VB.OptionButton DesativadosNao 
         BackColor       =   &H00EAE8DD&
         Caption         =   "Não Mostrar Desativados"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.OptionButton DesativadosSim 
         BackColor       =   &H00EAE8DD&
         Caption         =   "Somente Desativados"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   510
         Width           =   2295
      End
      Begin VB.OptionButton DesativadosTodos 
         BackColor       =   &H00EAE8DD&
         Caption         =   "Todos os Produtos"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   2295
      End
   End
   Begin VB.TextBox codigofor 
      Height          =   285
      Left            =   3120
      TabIndex        =   18
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox fornecedor 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   1440
      Width           =   6615
   End
   Begin Crystal.CrystalReport CryRelatorio 
      Left            =   4320
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.ComboBox Ordem 
      Height          =   315
      Left            =   240
      TabIndex        =   6
      Top             =   3480
      Width           =   2175
   End
   Begin VB.TextBox Copias 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2520
      TabIndex        =   7
      Text            =   "1"
      Top             =   3480
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAE8DD&
      Caption         =   "Tipo de Pesquisa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2040
      TabIndex        =   13
      Top             =   1800
      Width           =   2175
      Begin VB.OptionButton Igual 
         BackColor       =   &H00EAE8DD&
         Caption         =   "Igual a"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   1575
      End
      Begin VB.OptionButton Qualquer 
         BackColor       =   &H00EAE8DD&
         Caption         =   "Em Qualquer Parte"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   570
         Width           =   1695
      End
      Begin VB.OptionButton Iniciado 
         BackColor       =   &H00EAE8DD&
         Caption         =   "Iniciado por"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAE8DD&
      Caption         =   "Saída"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      TabIndex        =   10
      Top             =   1800
      Width           =   1695
      Begin VB.OptionButton Video 
         BackColor       =   &H00EAE8DD&
         Caption         =   "Vídeo"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Impressora 
         BackColor       =   &H00EAE8DD&
         Caption         =   "Impressora"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar  F10"
      Height          =   495
      Left            =   5400
      TabIndex        =   9
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Confirmar F3"
      Height          =   495
      Left            =   3600
      TabIndex        =   8
      Top             =   3360
      Width           =   1455
   End
   Begin VB.TextBox Nome 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   6615
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
      Left            =   240
      TabIndex        =   17
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Classificar por"
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
      Left            =   240
      TabIndex        =   15
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copias"
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
      Left            =   2520
      TabIndex        =   14
      Top             =   3240
      Width           =   750
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Relatório de Produto por Nome"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   120
      Width           =   6855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome"
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
      Left            =   240
      TabIndex        =   11
      Top             =   600
      Width           =   630
   End
End
Attribute VB_Name = "FrmRelProdutoNome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type TipoFor
      Codigo As String
      Nome As String
End Type
Private a As Integer
Private Mtfor() As TipoFor
Private LcTamanho As Long
Private Rs        As ADODB.Recordset
Private Rel       As New CryRelProdutoNome
Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then
   SendKeys "{TAB}"
Else
  If KeyCode = 116 And Index <> 7 Then
  Else
    Call Teclas(KeyCode)
  End If
End If
End Sub

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

StrSql = "SELECT relProdutos.Desativado,relProdutos.CodUsuario, relProdutos.NOME, relProdutos.CODBAR, relProdutos.Custo, relProdutos.Preco,relProdutos.LimiteVenda, relProdutos.MinimoVenda, relProdutos.MinimoEst, relProdutos.UnidMedida, relProdutos.QtdMedida, relProdutos.CST, relProdutos.lucro, relProdutos.ComissaoFornecedor, relProdutos.Fornecedor, relProdutos.codigo, relProdutos.QuantEstoque, relProdutos.ipi, relProdutos.percentualcusto, relProdutos.maximoEstoque, relProdutos.custoTotal, relProdutos.subitens, relProdutos.multiplositens, relProdutos.Santa1, relProdutos.santa2, relProdutos.California"
StrSql = StrSql & " FROM relProdutos order by NOME"
'Debug.Print StrSql
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
Screen.MousePointer = vbDefault
End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then
   SendKeys "{TAB}"
Else
  If KeyCode = 116 And Index <> 7 Then
  Else
    Call Teclas(KeyCode)
  End If
End If
End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then
   SendKeys "{TAB}"
Else
  If KeyCode = 116 And Index <> 7 Then
  Else
    Call Teclas(KeyCode)
  End If
End If
End Sub

Private Sub Copias_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Form_Activate()
On Error Resume Next
Set GlFormA = Me
End Sub
Function NomeImpressora() As String
On Error Resume Next
NomeImpressora = Rel.PrinterName
Relatorios.SetFocus
End Function
Function SetaImpressora() As String
On Error Resume Next

    Rel.PrinterSetup Relatorios.hWnd
    SetaImpressora = Rel.PrinterName

Relatorios.SetFocus
End Function
Function CarregaCombo()
Ordem.AddItem "Nome"
Ordem.AddItem "Código"
Ordem.AddItem "Preço"
Ordem.AddItem "Estoque"
Ordem.Text = "Nome"
End Function
Private Sub Form_Load()
On Error Resume Next
DataS.Text = Format(GlDataSistema, "dd/mm/yyyy")
HoraS.Text = Format(Time, "hh:mm:ss")
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
LcIndice = "CODIGO"
'Me.Height = 3705
'Me.Width = 7080
CarregaCombo
fornecedor.Visible = GlRepresentante
Label5.Visible = GlRepresentante
If GlRepresentante Then Carregaforn
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
FrmPrincipal.SetFocus
End Sub

Private Sub fornecedor_Click()
Dim LcAchou As Boolean
LcAchou = False
For a = 0 To LcTamanho
  If Mtfor(a).Nome = fornecedor.Text Then
     codigofor.Text = Mtfor(a).Codigo
     LcAchou = True
     Exit For
  End If
Next
If Not LcAchou Then codigofor.Text = ""
  
End Sub

Private Sub Igual_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then
   SendKeys "{TAB}"
Else
  If KeyCode = 116 And Index <> 7 Then
  Else
    Call Teclas(KeyCode)
  End If
End If
End Sub

Private Sub Impressora_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then
   SendKeys "{TAB}"
Else
  If KeyCode = 116 And Index <> 7 Then
  Else
    Call Teclas(KeyCode)
  End If
End If
End Sub

Private Sub Iniciado_Click()
'Escolha
'BuscaExpressao

End Sub

Private Sub Iniciado_GotFocus()
On Error Resume Next
'Txt(0).Text = ""
'Txt(1).Text = ""
End Sub

Private Sub Iniciado_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then
   SendKeys "{TAB}"
Else
  If KeyCode = 116 And Index <> 7 Then
  Else
    Call Teclas(KeyCode)
  End If
End If
End Sub

Private Sub Nome_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then
   SendKeys "{TAB}"
Else
  If KeyCode = 116 And Index <> 7 Then
  Else
    Call Teclas(KeyCode)
  End If
End If
End Sub

Private Sub Ordem_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Qualquer_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then
   SendKeys "{TAB}"
Else
  If KeyCode = 116 And Index <> 7 Then
  Else
    Call Teclas(KeyCode)
  End If
End If
End Sub

Private Sub Video_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then
   SendKeys "{TAB}"
Else
  If KeyCode = 116 And Index <> 7 Then
  Else
    Call Teclas(KeyCode)
  End If
End If
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
Function GeraDados()
On Error GoTo errGera
Dim RsUnidade   As Recordset
Dim LcUnidade   As String
Dim RsNota      As ADODB.Recordset
Dim RsNotaMdb   As ADODB.Recordset
Dim LcSql       As String
Dim LcNome      As String
Dim db          As Database

Set db = OpenDatabase(GLBase)
'==> Apagando Registros
afetados = ExecutaSql("Delete from relprodutos")
Dim LcWhereDesativado As String

If DesativadosSim.Value Then
  LcWhereDesativado = " And Desativado<>0"
End If
If DesativadosNao.Value Then
  LcWhereDesativado = " And Desativado=0"
End If

LcSql = "Select * from produtos where nome like '" & UCase(Nome.Text) & "%'" & LcWhereDesativado & " and Fornecedor like '" & UCase(codigofor.Text) & "%' order by codigo"
Set RsNota = AbreRecordset(LcSql, True)
Do Until RsNota.EOF
    LcSql = "SELECT SIMBOLO FROM ALID004 Where Cod='" & RsNota!UnidMedida & "'"
    Set RsUnidade = db.OpenRecordset(LcSql, dbOpenDynaset)
    If Not RsUnidade.EOF Then
        LcUnidade = RsUnidade!Simbolo & ""
    Else
        LcUnidade = ""
    End If
    RsUnidade.Close
    Set RsUnidade = Nothing
    If Not IsNull(RsNota!QuantEstoque) Then
       If Len(RsNota!QuantEstoque) > 0 Then
          LcEstoque = RsNota!QuantEstoque
       Else
          LcEstoque = 0
       End If
    Else
        LcEstoque = 0
    End If
    LcSql = " Insert Into relprodutos(codigo,NOME,Custo,MinimoVenda,MinimoEst,"
    LcSql = LcSql & "UnidMedida,QtdMedida,QuantEstoque,Preco,Desativado,LimiteVenda)Values("
    LcSql = LcSql & "'" & RsNota!Codigo & "','" & RsNota!Nome & "',"
    LcSql = LcSql & "" & Replace(RsNota!Custo, ",", ".") & "," & Replace(RsNota!MinimoVenda, ",", ".") & ","
    LcSql = LcSql & "" & Replace(RsNota!MinimoEst, ",", ".") & ",'" & LcUnidade & "',"
    LcSql = LcSql & "" & Replace(RsNota!QtdMedida, ",", ".") & "," & LcEstoque & ","
    LcSql = LcSql & "" & Replace(RsNota!Preco, ",", ".") & ","
    LcSql = LcSql & "" & IIf(RsNota!Desativado, 0, -1) & ","
    LcSql = LcSql & "" & Replace(RsNota!LimiteVenda, ",", ".") & ")"
    'Debug.Print LcSql
    afetados = ExecutaSql(LcSql)
  
    DoEvents
    RsNota.MoveNext
Loop
RsNota.Close
Exit Function
errGera:
'MsgBox err.Description & err.Number
Resume Next
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
lctitulo = "Relatorio de Produtos Por Nome"
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

