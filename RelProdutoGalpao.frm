VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form RelProdutoEstoque 
   BackColor       =   &H00EAE8DD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatorio de Estoque nos Galpões"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7080
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   7080
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00EAE8DD&
      Caption         =   "Desativados"
      Height          =   1335
      Left            =   4440
      TabIndex        =   15
      Top             =   1680
      Width           =   2535
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
      Begin VB.OptionButton DesativadosSim 
         BackColor       =   &H00EAE8DD&
         Caption         =   "Somente Desativados"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   510
         Width           =   2295
      End
      Begin VB.OptionButton DesativadosTodos 
         BackColor       =   &H00EAE8DD&
         Caption         =   "Todos os Produtos"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   2295
      End
   End
   Begin Crystal.CrystalReport CryRelatorio 
      Left            =   4680
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox Nome 
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   840
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Confirmar F3"
      Height          =   495
      Left            =   5160
      TabIndex        =   9
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar  F10"
      Height          =   495
      Left            =   5160
      TabIndex        =   8
      Top             =   1095
      Width           =   1455
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
      TabIndex        =   5
      Top             =   1680
      Width           =   1815
      Begin VB.OptionButton Impressora 
         BackColor       =   &H00EAE8DD&
         Caption         =   "Impressora"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton Video 
         BackColor       =   &H00EAE8DD&
         Caption         =   "Vídeo"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Value           =   -1  'True
         Width           =   1095
      End
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
      Left            =   2160
      TabIndex        =   1
      Top             =   1680
      Width           =   2175
      Begin VB.OptionButton Iniciado 
         BackColor       =   &H00EAE8DD&
         Caption         =   "Iniciado por"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton Qualquer 
         BackColor       =   &H00EAE8DD&
         Caption         =   "Em Qualquer Parte"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   570
         Width           =   1695
      End
      Begin VB.OptionButton Igual 
         BackColor       =   &H00EAE8DD&
         Caption         =   "Igual a"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   1575
      End
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
      Left            =   240
      TabIndex        =   0
      Text            =   "1"
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Produto"
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
      TabIndex        =   14
      Top             =   600
      Width           =   825
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Relação de Estoque dos Galpões"
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
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   6855
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
      Left            =   240
      TabIndex        =   12
      Top             =   3120
      Width           =   750
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Para selecionar todos os produtos apenas clique em Confirmar"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1320
      Width           =   4815
   End
End
Attribute VB_Name = "RelProdutoEstoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Rs As ADODB.Recordset
Private Rel As New CryRelProdutoGalpao
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

StrSql = "SELECT relProdutos.Desativado, relProdutos.CodUsuario, relProdutos.NOME, relProdutos.CODBAR, relProdutos.Custo, relProdutos.Preco, relProdutos.MinimoVenda, relProdutos.MinimoEst, relProdutos.UnidMedida, relProdutos.QtdMedida, relProdutos.CST, relProdutos.lucro, relProdutos.ComissaoFornecedor, relProdutos.Fornecedor, relProdutos.codigo, relProdutos.QuantEstoque, relProdutos.ipi, relProdutos.percentualcusto, relProdutos.maximoEstoque, relProdutos.custoTotal, relProdutos.subitens, relProdutos.multiplositens, relProdutos.Santa1, relProdutos.santa2, relProdutos.California"
StrSql = StrSql & " FROM relProdutos order by relProdutos.NOME"

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

LcSql = "Select Desativado,codigo,NOME, Santa1,Santa2,California"
LcSql = LcSql & " From produtos Where NOME like '" & Nome.Text & "%'" & LcWhereDesativado & " order by codigo"
Set RsNota = AbreRecordset(LcSql, True)
Do Until RsNota.EOF
    LcSql = " Insert Into relprodutos(codigo,NOME,Santa1,Santa2,California,Desativado"
    LcSql = LcSql & ")Values("
    LcSql = LcSql & "'" & RsNota!codigo & "','" & RsNota!Nome & "',"
    LcSql = LcSql & "" & Replace(RsNota!santa1, ",", ".") & "," & Replace(RsNota!Santa2, ",", ".") & ","
    LcSql = LcSql & Replace(RsNota!california, ",", ".") & ","
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



Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Command2_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Copias_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Form_Activate()
Set GlFormA = Me
End Sub

Private Sub Igual_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Impressora_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
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
