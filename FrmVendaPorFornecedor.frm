VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form FrmVendaPorFornecedor 
   BackColor       =   &H00CBB19C&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vendas por Fornecedor"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   8895
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Saída"
      Height          =   1335
      Left            =   240
      TabIndex        =   17
      Top             =   2640
      Width           =   1815
      Begin VB.OptionButton Video 
         Caption         =   "Video"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton impressora 
         Caption         =   "Impressora"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.TextBox copias 
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
      Left            =   4920
      TabIndex        =   16
      Text            =   "1"
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Confirma F3"
      Height          =   615
      Left            =   6840
      TabIndex        =   15
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar"
      Height          =   615
      Left            =   6840
      TabIndex        =   14
      Top             =   840
      Width           =   1695
   End
   Begin VB.ComboBox Fornecedor 
      Height          =   315
      ItemData        =   "FrmVendaPorFornecedor.frx":0000
      Left            =   240
      List            =   "FrmVendaPorFornecedor.frx":0002
      TabIndex        =   13
      Top             =   480
      Width           =   4575
   End
   Begin VB.TextBox codigo 
      Height          =   405
      Left            =   2760
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo "
      Height          =   1335
      Left            =   2160
      TabIndex        =   9
      Top             =   2640
      Width           =   2175
      Begin VB.OptionButton analitico 
         Caption         =   "Analítico"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton sintetico 
         Caption         =   "Sintético"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Situção"
      Height          =   1335
      Left            =   4440
      TabIndex        =   5
      Top             =   2640
      Width           =   1695
      Begin VB.OptionButton pago 
         Caption         =   "Pago"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton naoPago 
         Caption         =   "Não Pago"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton todos 
         Caption         =   "Todas"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Apos Imprimir"
      Height          =   1335
      Left            =   6360
      TabIndex        =   2
      Top             =   2640
      Visible         =   0   'False
      Width           =   2295
      Begin VB.OptionButton marcarpg 
         Caption         =   "Marcar como Pago"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton marcar 
         Caption         =   "Não Marcar como Pago"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.ComboBox vendedor 
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   4575
   End
   Begin VB.TextBox codigoVendedor 
      Height          =   405
      Left            =   4440
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin Crystal.CrystalReport CryRelatorio 
      Left            =   5280
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSMask.MaskEdBox Datai 
      Height          =   375
      Left            =   240
      TabIndex        =   20
      Top             =   2160
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yy"
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Dataf 
      Height          =   375
      Left            =   2160
      TabIndex        =   21
      Top             =   2160
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yy"
      Mask            =   "99/99/99"
      PromptChar      =   " "
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
      Height          =   240
      Left            =   4920
      TabIndex        =   26
      Top             =   120
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Line Line1 
      X1              =   6240
      X2              =   6240
      Y1              =   -240
      Y2              =   1800
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
      Height          =   240
      Index           =   0
      Left            =   240
      TabIndex        =   25
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Final"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2160
      TabIndex        =   24
      Top             =   1800
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Inicial"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   240
      TabIndex        =   23
      Top             =   1800
      Width           =   1185
   End
   Begin VB.Line Line2 
      X1              =   6240
      X2              =   8760
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vendedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   240
      TabIndex        =   22
      Top             =   960
      Width           =   1035
   End
End
Attribute VB_Name = "FrmVendaPorFornecedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type TipoVend
      codigo As String
      Nome As String
End Type
Private Rel As New CrysPropostaVenda
Private MtMatVendedor() As TipoVend
Private MtVendedor() As TipoVend

Function AbreRecordsetRel(LcSql As String, RsAtual As ADODB.Recordset) As ADODB.Recordset

On Error GoTo ErroAbreRs
LcComentario = "- AbreRecordset - Criando Nova Instancia do RecordSet."
Set RsAtual = New ADODB.Recordset
LcComentario = "- AbreRecordset - Setando os Parametros do Recordset."
RsAtual.CursorType = adOpenDynamic ' adOpenStatic
RsAtual.CursorLocation = adUseClient
RsAtual.LockType = adLockReadOnly
RsAtual.Source = LcSql
RsAtual.ActiveConnection = conexaoAdo

LcComentario = "- AbreRecordset - Abrindo o Recordset."
RsAtual.Open
Set AbreRecordsetRel = RsAtual
Exit Function

ErroAbreRs:
'If err.Number = 3709 Then
'   'abreconexao
'   Resume 0
'End If
'If LcExibemsg Then ErrosSistema = MsgBox(msg, 64, "erro Abrindo Tabela. ") Else ErrosSistema = 0
'MsgBox err.Description & err.Number
'Resume 0
logErro err.Number, err.Description, LcComentario
Resume Next
End Function
Function CarregaVendedor()

On Error GoTo errc
Dim RsVendedor As Recordset, RsVend As Recordset
AbreBase
Set RsVendedor = Dbbase.OpenRecordset("ALID002", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsVend = Dbbase.OpenRecordset("Alid200", dbOpenDynaset, dbSeeChanges, dbOptimistic)

LcTamanho = 0
Do Until RsVendedor.EOF
   If Not IsNull(RsVendedor!RazaoSoc) Then
      ReDim Preserve MtVendedor(LcTamanho)
      MtVendedor(LcTamanho).codigo = RsVendedor!codigo
      MtVendedor(LcTamanho).Nome = RsVendedor!RazaoSoc & ""
      Fornecedor.AddItem RsVendedor!RazaoSoc
      LcTamanho = LcTamanho + 1
   End If
   RsVendedor.MoveNext
  
Loop
If LcTamanho > 0 Then LcTamanho = LcTamanho - 1
LcTamMatVend = 0
Do Until RsVend.EOF
  If Not IsNull(RsVend!Nome) Then
    ReDim Preserve MtMatVendedor(LcTamMatVend)
    MtMatVendedor(LcTamMatVend).codigo = RsVend!codigo
    MtMatVendedor(LcTamMatVend).Nome = RsVend!Nome & ""
    vendedor.AddItem RsVend!Nome
    LcTamMatVend = LcTamMatVend + 1
  End If
  RsVend.MoveNext
   
Loop
 If LcTamMatVend > 0 Then LcTamMatVend = LcTamMatVend - 1
'Comissao.AddItem "TODOS"
'Comissao.Text = "TODOS"
RsVendedor.Close
Set RsVendedor = Nothing
RsVend.Close
Set RsVend = Nothing
Exit Function
errc:

Exit Function

End Function

Private Sub Command1_Click()
Dim LcSql As String
Dim StrWhere As String
Dim LcCodigoFor As String
Dim LcCodigoVend As String

Dim Rs As ADODB.Recordset

If Len(Fornecedor.Text) > 0 Then
    For a = 0 To LcTamanho
        If MtVendedor(a).Nome = Fornecedor.Text Then
           LcCodigoFor = MtVendedor(a).codigo
           Exit For
        End If
    Next
End If


If Len(vendedor.Text) > 0 Then
    For a = 0 To LcTamanho
        If MtMatVendedor(a).Nome = vendedor.Text Then
           LcCodigoVend = MtMatVendedor(a).codigo
           Exit For
        End If
    Next
End If
If Len(LcCodigoFor) > 0 Then
   StrWhere = "produtos.Fornecedor)='" & LcCodigoFor & "'"
End If
If Len(LcCodigoVend) > 0 Then
   If Len(StrWhere) > 0 Then StrWhere = StrWhere & " and "
   StrWhere = StrWhere & "alid050.Vendedor)='" & LcCodigoVend & "'"
End If


LcSql = "SELECT alid050.NUMNF, alid050.DTEMIS, alid050.Vendedor, alid050.NomeVendedorImprimir, alid050.finalidadeEmissao, alid050.CLIENTE, alid050.NomeCliente, alid052.codProd, alid052.descricao, alid052.UNIMED, alid052.QTDE, alid052.VALUNIT, alid052.desconto, produtos.codigo, produtos.Fornecedor, (alid052!VALUNIT*alid052!QTDE)-alid052!desconto AS VALORTOTAL"
LcSql = LcSql & " FROM (alid050 INNER JOIN alid052 ON alid050.codigo=alid052.CodigoNota) INNER JOIN produtos ON alid052.codProd=produtos.codigo "
 
If Len(StrWhere) > 0 Then LcSql = LcSql & " Where " & StrWhere

Set Rs = AbreRecordsetRel(LcSql, Rs)

Load Relatorios
    With Relatorios
         Rel.DiscardSavedData
         Rel.Database.SetDataSource Rs
         .CRViewer1.ReportSource = Rel
    End With
 setaformula
Relatorios.CRViewer1.ViewReport
    Relatorios.Show

Screen.MousePointer = vbDefault
End Sub
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

If IsDate(Format(Datai.Text, "dd/mm/yy")) And IsDate(Format(Dataf.Text, "dd/mm/yy")) Then
   lctitulo = "Relatorio de Notas de Saida: " & Datai.Text & " à " & Dataf.Text
   Else
   lctitulo = "Relatorio de Notas de Saida"
End If
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
Private Sub Form_Load()
CarregaVendedor
End Sub
