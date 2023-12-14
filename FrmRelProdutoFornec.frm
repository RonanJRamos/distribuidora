VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form FrmRelProdutoFornec 
   BackColor       =   &H00EAE8DD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Produtos por Fornecedor"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6960
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   6960
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H00EAE8DD&
      Caption         =   "Desativados"
      Height          =   1335
      Left            =   4080
      TabIndex        =   17
      Top             =   2040
      Width           =   2535
      Begin VB.OptionButton DesativadosTodos 
         BackColor       =   &H00EAE8DD&
         Caption         =   "Todos os Produtos"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   2295
      End
      Begin VB.OptionButton DesativadosSim 
         BackColor       =   &H00EAE8DD&
         Caption         =   "Somente Desativados"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   510
         Width           =   2295
      End
      Begin VB.OptionButton DesativadosNao 
         BackColor       =   &H00EAE8DD&
         Caption         =   "Não Mostrar Desativados"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Value           =   -1  'True
         Width           =   2295
      End
   End
   Begin VB.CheckBox SemCusto 
      BackColor       =   &H00EAE8DD&
      Caption         =   "Sem Custo"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1560
      Width           =   2655
   End
   Begin Crystal.CrystalReport CryRelatorio 
      Left            =   4800
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox Nome 
      Height          =   375
      Left            =   3480
      TabIndex        =   15
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox Fornecedor 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   4575
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
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Text            =   "1"
      Top             =   3600
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
      Left            =   1920
      TabIndex        =   12
      Top             =   2040
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
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   1815
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
      Left            =   5280
      TabIndex        =   8
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Confirmar F3"
      Height          =   495
      Left            =   5280
      TabIndex        =   7
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Para selecionar todos os produtos apenas clique em Confirmar"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   1320
      Width           =   4815
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
      Left            =   120
      TabIndex        =   13
      Top             =   3360
      Width           =   750
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Relatório de produtos por Fornecedores"
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
      TabIndex        =   10
      Top             =   120
      Width           =   6855
   End
   Begin VB.Label Label1 
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
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "FrmRelProdutoFornec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type DadoFornecedor
        Codigo As String
        Nome As String
End Type
Private MtFornecedor() As DadoFornecedor
Private LcTam, a As Long
Private Rs        As ADODB.Recordset
Private Rel       As New CryRelProdutoFornecedor
Function SetaImpressora() As String
On Error Resume Next
    Rel.PrinterSetup Relatorios.hWnd
    SetaImpressora = Rel.PrinterName
    Relatorios.SetFocus
End Function
Private Sub Command1_Click()
On Error Resume Next
Dim StrSql          As String
If Len(Nome.Text) = 0 Then
    MsgBox "Escolha um Fornecedor para Visualizar o Relatorio.", 64, "Aviso"
    Fornecedor.SetFocus
    Exit Sub
End If
LcCap = Me.Caption
Me.Caption = "Aguarde, processando os dados..."
Screen.MousePointer = 11
GeraDados
Screen.MousePointer = 0
Me.Caption = LcCap
Screen.MousePointer = vbHourglass
If SemCusto.Value = 0 Then
    StrSql = "SELECT relProdutos.Desativado,relProdutos.CodUsuario, relProdutos.NOME, relProdutos.CODBAR, relProdutos.Custo, relProdutos.Preco, relProdutos.MinimoVenda, relProdutos.MinimoEst, relProdutos.UnidMedida, relProdutos.QtdMedida, relProdutos.CST, relProdutos.lucro, relProdutos.ComissaoFornecedor, relProdutos.Fornecedor, relProdutos.codigo, relProdutos.QuantEstoque, relProdutos.ipi, relProdutos.percentualcusto, relProdutos.maximoEstoque, relProdutos.custoTotal, relProdutos.subitens, relProdutos.multiplositens, relProdutos.Santa1, relProdutos.santa2, relProdutos.California"
    StrSql = StrSql & " FROM relProdutos order by relProdutos.NOME"

Else
    StrSql = "SELECT relProdutos.Desativado,relProdutos.CodUsuario, relProdutos.NOME, relProdutos.CODBAR, 0 as Custo, relProdutos.Preco, relProdutos.MinimoVenda, relProdutos.MinimoEst, relProdutos.UnidMedida, relProdutos.QtdMedida, relProdutos.CST, relProdutos.lucro, relProdutos.ComissaoFornecedor, relProdutos.Fornecedor, relProdutos.codigo, relProdutos.QuantEstoque, relProdutos.ipi, relProdutos.percentualcusto, relProdutos.maximoEstoque, relProdutos.custoTotal, relProdutos.subitens, relProdutos.multiplositens, relProdutos.Santa1, relProdutos.santa2, relProdutos.California"
    StrSql = StrSql & " FROM relProdutos order by relProdutos.NOME"
End If

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
End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Copias_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Form_Activate()
Set GlFormA = Me
End Sub

Private Sub Form_Load()
On Error Resume Next
DataS.Text = Format(GlDataSistema, "dd/mm/yyyy")
HoraS.Text = Format(Time, "hh:mm:ss")
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
LcIndice = "CODIGO"
Me.Height = 4000
Me.Width = 7080
carregaFornecedor
End Sub
Function carregaFornecedor()
Dim LcEmpresa, LcEndereco, LcFone, LcVer, LcCap, LcVer1, LcSql As String
Dim RsEmpresa As Recordset
AbreBase
LcTam = 0
LcSql = "Select * from alid002 order by razaosoc"
Set RsEmpresa = Dbbase.OpenRecordset(LcSql, dbOpenDynaset) ', dbSeeChanges, dbOptimistic)
Do Until RsEmpresa.EOF
    ReDim Preserve MtFornecedor(LcTam)
    If Not IsNull(RsEmpresa!RazaoSoc) Then
        MtFornecedor(LcTam).Codigo = RsEmpresa!Codigo
        MtFornecedor(LcTam).Nome = RsEmpresa!RazaoSoc
        Fornecedor.AddItem RsEmpresa!RazaoSoc
        LcTam = LcTam + 1
    End If
    RsEmpresa.MoveNext
Loop
If LcTam > 0 Then LcTam = LcTam - 1
RsEmpresa.Close
Dbbase.Close
Set RsEmpresa = Nothing
Set dbbasee = Nothing

End Function


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
FrmPrincipal.SetFocus
End Sub

Private Sub fornecedor_Click()
For a = 0 To LcTam
    If MtFornecedor(a).Nome = Fornecedor.Text Then
       Nome.Text = MtFornecedor(a).Codigo
       Exit For
    End If
Next

End Sub

Private Sub Fornecedor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Igual_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Impressora_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
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
End Sub

Private Sub Nome_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Qualquer_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Video_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub
Function GeraDados()
On Error GoTo errGera
Dim RsUnidade   As Recordset
Dim LcUnidade   As String
Dim RsNota      As ADODB.Recordset
Dim RsNotaMdb   As ADODB.Recordset
Dim LcSql       As String
Dim LcNome      As String
Dim LcFornecedor As String
Dim db          As Database
Dim LcWhereDesativado As String

If DesativadosSim.Value Then
  LcWhereDesativado = " And Desativado<>0"
End If
If DesativadosNao.Value Then
  LcWhereDesativado = " And Desativado=0"
End If


Set db = OpenDatabase(GLBase)
'==> Apagando Registros
afetados = ExecutaSql("Delete from relprodutos")

LcSql = "Select * from produtos where Fornecedor='" & UCase(Nome.Text) & "'" & LcWhereDesativado & " order by codigo"
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
    
    LcSql = "SELECT RAZAOSOC FROM ALID002 Where Codigo='" & RsNota!Fornecedor & "'"
    Set RsUnidade = db.OpenRecordset(LcSql, dbOpenDynaset)
    If Not RsUnidade.EOF Then
        LcFornecedor = RsUnidade!RazaoSoc & ""
    Else
        LcFornecedor = ""
    End If
    RsUnidade.Close
    Set RsUnidade = Nothing
    Dim LcEstoque As Double
    If RsNota!QtdMedida > 0 Then
        LcEstoque = RsNota!QuantEstoque / RsNota!QtdMedida
    Else
        LcEstoque = RsNota!QuantEstoque
    End If
    
    LcSql = " Insert Into relprodutos(codigo,NOME,Custo,MinimoVenda,MinimoEst,"
    LcSql = LcSql & "UnidMedida,QtdMedida,QuantEstoque,Preco,Fornecedor,Desativado)Values("
    LcSql = LcSql & "'" & RsNota!Codigo & "','" & RsNota!Nome & "',"
    LcSql = LcSql & "" & Replace(RsNota!Custo, ",", ".") & "," & Replace(RsNota!MinimoVenda, ",", ".") & ","
    LcSql = LcSql & "" & Replace(RsNota!MinimoEst, ",", ".") & ",'" & LcUnidade & "',"
    LcSql = LcSql & "" & Replace(RsNota!QtdMedida, ",", ".") & "," & Replace(LcEstoque, ",", ".") & ","
    LcSql = LcSql & "" & Replace(RsNota!Preco, ",", ".") & ",'" & LcFornecedor & "',"
    LcSql = LcSql & "" & IIf(RsNota!Desativado, 0, -1) & ")"
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
lctitulo = "Relatorio de Produtos (Por Fornecedor)"
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

