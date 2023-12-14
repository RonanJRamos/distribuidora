VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form comprasfornecedor 
   BackColor       =   &H00EAE8DD&
   Caption         =   "Relatorio de Compras por Fornecedor"
   ClientHeight    =   2805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   ScaleHeight     =   2805
   ScaleWidth      =   4230
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00EAE8DD&
      Caption         =   "Desativados"
      Height          =   1335
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   2535
      Begin VB.OptionButton DesativadosNao 
         BackColor       =   &H00EAE8DD&
         Caption         =   "Não Mostrar Desativados"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.OptionButton DesativadosSim 
         BackColor       =   &H00EAE8DD&
         Caption         =   "Somente Desativados"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   510
         Width           =   2295
      End
      Begin VB.OptionButton DesativadosTodos 
         BackColor       =   &H00EAE8DD&
         Caption         =   "Todos os Produtos"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   2295
      End
   End
   Begin Crystal.CrystalReport CryRelatorio 
      Left            =   3960
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox codigo 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar"
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Visualizar"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   1575
   End
   Begin VB.ComboBox fornec 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fornecedor"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   810
   End
End
Attribute VB_Name = "comprasfornecedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type forn
    Nome As String
    codigo As String
End Type
Private LcMat() As forn
Private LcRegs As Long

Private Rs        As ADODB.Recordset
Private Rel       As New CryRelProdutoCompras
Function SetaImpressora() As String
On Error Resume Next
    Rel.PrinterSetup Relatorios.hWnd
    SetaImpressora = Rel.PrinterName
    Relatorios.SetFocus
End Function
Private Sub Command1_Click()
On Error Resume Next
Dim StrSql          As String
If Len(codigo.Text) = 0 Then
    MsgBox "Escolha um Fornecedor para Visualizar o Relatorio.", 64, "Aviso"
    fornec.SetFocus
    Exit Sub
End If
LcCap = Me.Caption
Me.Caption = "Aguarde, processando os dados..."
Screen.MousePointer = 11
GeraDados
Screen.MousePointer = 0
Me.Caption = LcCap
Screen.MousePointer = vbHourglass

StrSql = "SELECT relProdutos.Desativado,relProdutos.CodUsuario, relProdutos.NOME, relProdutos.CODBAR, relProdutos.Custo, relProdutos.Preco, relProdutos.MinimoVenda, relProdutos.MinimoEst, relProdutos.UnidMedida, relProdutos.QtdMedida, relProdutos.CST, relProdutos.lucro, relProdutos.ComissaoFornecedor, relProdutos.Fornecedor, relProdutos.codigo, relProdutos.QuantEstoque, relProdutos.ipi, relProdutos.percentualcusto, relProdutos.maximoEstoque, relProdutos.custoTotal, relProdutos.subitens, relProdutos.multiplositens, relProdutos.Santa1, relProdutos.santa2, relProdutos.California"
StrSql = StrSql & " FROM relProdutos"

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
Private Sub Command2_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub Form_Activate()
Set GlFormA = Me
End Sub

Private Sub Form_Load()
Carregaforn
End Sub
Function Carregaforn()
Dim Rs As Recordset
AbreBase
LcRegs = 0
Set Rs = Dbbase.OpenRecordset("Select * from alid002 order by razaosoc", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Do Until Rs.EOF
   ReDim Preserve LcMat(LcRegs)
   LcMat(LcRegs).Nome = Rs!RAZAOSOC & ""
   LcMat(LcRegs).codigo = Rs!codigo & ""
   fornec.AddItem Rs!RAZAOSOC & ""
   Rs.MoveNext
   LcRegs = LcRegs + 1
Loop
LcRegs = LcRegs - 1
End Function

Private Sub fornec_Click()
Dim a As Long
For a = 0 To LcRegs
    If UCase(LcMat(a).Nome) = UCase(fornec.Text) Then
       codigo.Text = LcMat(a).codigo
       Exit For
    End If
Next

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
Dim LcFornecedor As String
Set db = OpenDatabase(GLBase)

Dim LcWhereDesativado As String

If DesativadosSim.Value Then
  LcWhereDesativado = " And Desativado<>0"
End If
If DesativadosNao.Value Then
  LcWhereDesativado = " And Desativado=0"
End If

'==> Apagando Registros
afetados = ExecutaSql("Delete from relprodutos")

LcSql = "Select Desativado,codigo,NOME, MinimoEst,QuantEstoque,QtdMedida ,Fornecedor"
LcSql = LcSql & " From produtos Where Fornecedor='" & codigo.Text & "'" & LcWhereDesativado & " order by codigo"
Set RsNota = AbreRecordset(LcSql, True)
Do Until RsNota.EOF
    LcSql = "SELECT RAZAOSOC FROM ALID002 Where Codigo='" & RsNota!Fornecedor & "'"
    Set RsUnidade = db.OpenRecordset(LcSql, dbOpenDynaset)
    If Not RsUnidade.EOF Then
        LcFornecedor = RsUnidade!RAZAOSOC & ""
    Else
        LcFornecedor = ""
    End If
    RsUnidade.Close
    Set RsUnidade = Nothing
    LcSql = " Insert Into relprodutos(codigo,NOME,MinimoEst,QuantEstoque,Preco, Fornecedor,Desativado"
    LcSql = LcSql & ")Values("
    LcSql = LcSql & "'" & RsNota!codigo & "','" & RsNota!Nome & "',"
    LcSql = LcSql & "" & Replace(RsNota!MinimoEst, ",", ".") & "," & Replace(RsNota!QuantEstoque, ",", ".") & ","
    If CDbl(RsNota!QtdMedida) = 0 Then
        LcSql = LcSql & Replace(RsNota!QuantEstoque, ",", ".") & ","
    Else
        LcSql = LcSql & Replace(CDbl(RsNota!QuantEstoque) / CDbl(RsNota!QtdMedida), ",", ".") & ","
    End If
    LcSql = LcSql & "'" & LcFornecedor & "',"
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
lctitulo = "Relatorio de Produtos (Compras Por Fornecedor)"
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
