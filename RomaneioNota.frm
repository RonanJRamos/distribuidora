VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form RomaneioNota 
   BackColor       =   &H00D1CCAD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Romaneio por Nota Fiscal"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdFiltrar 
      Caption         =   "Filtrar"
      Height          =   255
      Left            =   3120
      TabIndex        =   8
      Top             =   360
      Width           =   1095
   End
   Begin MSMask.MaskEdBox dataf 
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Datai 
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin Crystal.CrystalReport CryRelatorio 
      Left            =   6480
      Top             =   6960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Imprimir / Gravar"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   7080
      Width           =   2895
   End
   Begin VB.CommandButton Gera 
      Caption         =   "->"
      Enabled         =   0   'False
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
      Left            =   5640
      TabIndex        =   3
      Top             =   3000
      Width           =   375
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   6855
      Left            =   6000
      TabIndex        =   2
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   12091
      _Version        =   393216
      Rows            =   1
      Cols            =   5
      FixedCols       =   0
      MergeCells      =   3
      AllowUserResizing=   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Fechar"
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   7080
      Width           =   2895
   End
   Begin MSComctlLib.TreeView acesso 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   11033
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   540
   End
End
Attribute VB_Name = "RomaneioNota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function GeraTree()
Dim Rs As ADODB.Recordset
Dim rsCliente As Recordset
Dim StrSql As String
Dim nodX As MSComctlLib.Node
Dim Item As String
Dim NF As String
Dim Cab As String
Dim No As String
Dim a As Integer
AbreBase
On Error GoTo errogera
StrSql = "SELECT alid050.romaneiogerado, alid050.NUMNF, alid050.DTEMIS, alid050.NATUREZA, alid050.CLIENTE, alid050.TRANSP, alid050.TIPOTRANS, alid050.PLACATRANS, alid050.UFTRANS, alid050.CGCCPFTRAN, alid050.ENDTRANS, alid050.MUNICTRANS, alid050.UFMUNIC, alid050.INSCEST, alid050.OBS02, alid050.OBS03, alid050.OBS04, alid050.CONTR, alid050.codigo, alid050.valorproduto, alid050.ValorNota, alid050.Vendedor, alid050.fonetransp, alid050.cidade, alid050.cep, alid050.formapag, alid050.dias, alid050.vencimento1, alid050.vencimento2, alid050.vencimento3, alid050.vencimento4, alid050.vencimento5, alid050.CondPag, alid050.status, alid050.DESCONTO, alid050.ICMS, alid050.CFOP, alid050.BaseIcms, alid050.ValorIcms, alid050.LocalEntrega, alid050.Oc, alid050.EnderecoEntrega, alid052.ITEM, alid052.GALPAO, alid052.QTDE, alid052.VALUNIT, alid052.CONTR, alid052.UNIMED, alid052.QTDUM, alid052.QTDE01, alid052.QTDE02, alid052.QTDE03, alid052.codigo, alid052.descricao, alid052.codProd, alid052.emissao, alid052.CST, alid052.icms " & _
          "FROM alid050 INNER JOIN alid052 ON alid050.NUMNF = alid052.NUMNF " & _
          "WHERE (((alid050.romaneiogerado)=0) and (alid050.DTEMIS Between '" & Format(DataI.Text, "yyyy-mm-dd") & "' And '" & Format(DataF.Text, "yyyy-mm-dd") & "' ));"

'Set Rs = Dbbase.OpenRecordset(StrSql)
'Debug.Print StrSql
Set Rs = AbreRecordset(StrSql)
'MsgBox DEscricaoErro

acesso.Nodes.Clear
a = 1
'exibe linhas
acesso.LineStyle = tvwTreeLines
'Inclui itens
Set nodX = acesso.Nodes.Add(, , "menu", "NOTAS FISCAIS disponiveis para Romaneio")
acesso.Nodes(a).ForeColor = &HC0&

nodX.Expanded = True
Do Until Rs.EOF
    StrSql = "Select * from alid001 where codigo='" & Rs!Cliente & "'"
    Set rsCliente = Dbbase.OpenRecordset(StrSql)
    
    Item = Right("000" & Rs!Qtde, 3) & " " & Rs!UNIMED & " C/ " & Rs!QTDUM & " - " & Rs!Descricao
    No = Rs!Item & "-" & Rs!NUMNF
    If NF <> Rs!NUMNF Then
       NF = Rs!NUMNF
       Cab = Rs!NUMNF & " - " & rsCliente!razaosoc
       Set nodX = acesso.Nodes.Add("menu", tvwChild, "Nf-" & NF, Cab)
       a = a + 1
       acesso.Nodes(a).ForeColor = 16711680
       Set nodX = acesso.Nodes.Add("Nf-" & NF, tvwChild, No, Item)
       'nodX.Expanded = True
       a = a + 1
    Else
       Set nodX = acesso.Nodes.Add("Nf-" & NF, tvwChild, No, Item)
       a = a + 1
       'nodX.Expanded = True
    End If
    Rs.MoveNext
Loop
Set Rs = Nothing

Exit Function
errogera:
If err = 35602 Then
   MsgBox "foi encontrado itens duplicados na NF:" & NF & Chr(13) & ". Por favor verifique a NF.", 64, "Erro Carregando Romaneio."
   Exit Function
Else
   MsgBox "foi encontrado um erro no pedido:" & NF & Chr(13) & err.Description & Chr(13) & ". Por favor verifique a NF.", 64, "Erro nº:" & err.Number
   Exit Function
End If
'MsgBox err.Number & err.Description
Resume 0

End Function

Private Sub acesso_Click()
On Error Resume Next
Gera.Enabled = True
End Sub

Private Sub CmdFiltrar_Click()
GeraTree
End Sub

Private Sub Command1_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub Command2_Click()
Dim Nume As Long
On Error Resume Next
LcCap = Me.Caption
Me.Caption = "Aguarde, Gravando o Romaneio..."
Screen.MousePointer = 11
Nume = GravaRomneio
Imprime Nume
MsgBox "Romaneio Salvo com Sucesso." & Chr(13) & "Com o Número:" & Nume, 64, "Aviso"
GeraTree
grid.Rows = 1
GeraGrid
Screen.MousePointer = 0
Me.Caption = LcCap

End Sub

Private Sub Form_Load()

GeraGrid
End Sub
Function GeraGrid()
grid.TextMatrix(0, 0) = "NF"
grid.TextMatrix(0, 1) = "Produto"
grid.TextMatrix(0, 2) = "Quant"
grid.TextMatrix(0, 3) = "Unidade"

grid.ColWidth(0) = "1000"
grid.ColWidth(1) = "3000"
grid.ColWidth(2) = "800"
grid.ColWidth(3) = "800"
grid.ColWidth(4) = "0"
End Function

Private Sub Gera_Click()
Dim LcStr As String
Dim a As Integer
Dim st() As String
Dim texto As String
Dim t() As String
 
LcStr = acesso.SelectedItem.Text
For a = 1 To acesso.Nodes.Count
    If LcStr = acesso.Nodes(a).Text Then Exit For
Next
texto = acesso.Nodes(a).FullPath
t = Split(texto, "\")

For b = 1 To acesso.Nodes.Count
 If acesso.Nodes(b).Text = t(1) Then
    If acesso.Nodes(b).ForeColor = &HC000& Then
       MsgBox "Esta NF ja foi incluida no romaneio.", 64, "Aviso"
       Exit Sub
    End If
    Exit For
 End If
Next
st = Split(acesso.Nodes(a).Key, "-")


buscaitens st(1)
For a = 1 To acesso.Nodes.Count
   If acesso.Nodes(a).Text = t(1) Then
      acesso.Nodes(a).ForeColor = &HC000&
      Exit For
   End If
Next
End Sub

Function buscaitens(LcNumero As String)
Dim Rs As ADODB.Recordset
Dim RsNota As ADODB.Recordset

Dim a As Integer
Dim Achou As Boolean
Dim quant As Double
Dim StrSql As String
Dim LcCodNota As String
Dim LcCap As String
LcCap = Me.Caption
Me.Caption = "Aguarde, Buscando itens da NF."
DoEvents
AbreBase
'Set Rs = Dbbase.OpenRecordset("Select * from alid052 where numnf='" & LcNumero & "' order by item")
'==> Busca o Codigo da nota
LcCodNota = 0
StrSql = "Select * from alid050 where (alid050.DTEMIS Between '" & Format(DataI.Text, "yyyy-mm-dd") & "' And '" & Format(DataF.Text, "yyyy-mm-dd") & "' )and numNf='" & LcNumero & "';"
Set RsNota = AbreRecordset(StrSql)
If Not RsNota.EOF Then
   LcCodNota = RsNota!Codigo
   LcNumNf = RsNota!NUMNF
End If
RsNota.Close
StrSql = "Select * from alid052 where CodigoNota=" & LcCodNota & " order by item"
Set Rs = AbreRecordset(StrSql, True)
If Rs.EOF Then
   Rs.Close
   StrSql = "Select * from alid052 where numnf='" & LcNumNf & "'"
   Set Rs = AbreRecordset(StrSql, True)
End If
Do Until Rs.EOF
  '==> Verifica se ja existe.
  Achou = False
  For a = 1 To grid.Rows - 1
     If grid.TextMatrix(a, 4) = Rs!codProd Then
        Achou = True
        Exit For
     End If
  Next
  If Achou Then
     If UCase(Rs!UNIMED & " C/ " & Rs!QTDUM) = UCase(grid.TextMatrix(a, 3)) Then
        quant = CDbl(grid.TextMatrix(a, 2)) + CDbl(Rs!Qtde)
        grid.TextMatrix(a, 2) = quant
        grid.TextMatrix(a, 0) = grid.TextMatrix(a, 0) & "-" & LcNumero
     Else
        '==> Converte para a unidade basica
        '==> Primeiro a do Grid
        GridSplit = Split(grid.TextMatrix(a, 3), "/")
        quantgrid = CDbl(GridSplit(1))
        quantgrid = quantgrid * CDbl(grid.TextMatrix(a, 2))
        '==>Agora, Converte a do Banco
        quant = CDbl(Rs!QTDUM) * CDbl(Rs!Qtde)
        quant = quant + quantgrid
        
        grid.TextMatrix(a, 2) = quant
        grid.TextMatrix(a, 0) = grid.TextMatrix(a, 0) & "-" & LcNumero
        
        grid.TextMatrix(a, 3) = "UN C/ 1"
     End If
  Else
     grid.Rows = grid.Rows + 1
     a = grid.Rows - 1
     grid.TextMatrix(a, 0) = LcNumero
     grid.TextMatrix(a, 1) = Rs!Descricao
     grid.TextMatrix(a, 2) = Rs!Qtde
     grid.TextMatrix(a, 3) = Rs!UNIMED & " C/ " & Rs!QTDUM
     grid.TextMatrix(a, 4) = Rs!codProd
  End If
  Rs.MoveNext
Loop
Set Rs = Nothing
Gera.Enabled = False
Me.Caption = LcCap

End Function
Function GravaRomneio() As String
Dim Rs As Recordset
Dim StrSql As String
Dim Numero As Long
Dim LcNf As String
AbreBase

StrSql = "Insert Into Romaneio (Data,Pedidos) Values (#"
StrSql = StrSql & Format(Date, "mm/dd/yy") & "#,'"

'==> Busca os Pedidos
For b = 1 To acesso.Nodes.Count
    If acesso.Nodes(b).ForeColor = &HC000& Then
        texto = acesso.Nodes(b).Key
        sp = Split(texto, "-")
        NPedidos = NPedidos & "-" & sp(1)
    End If
Next
StrSql = StrSql & NPedidos & "')"
Dbbase.Execute StrSql ', afetados
'If afetados = 1 Then
   Set Rs = Dbbase.OpenRecordset("Select * from Romaneio order by codigo desc")
   If Not Rs.EOF Then
      Numero = Rs!Codigo
   End If
   Set Rs = Nothing
'End If
For a = 1 To grid.Rows - 1
    StrSql = "Insert into DadosRomaneio (CodigoRomaneio,Pedidos,Produto,Quantidade,Unidade,CodigoProduto) Values ("
    StrSql = StrSql & Numero & ",'"
    StrSql = StrSql & grid.TextMatrix(a, 0) & "','"
    StrSql = StrSql & Replace(Replace(grid.TextMatrix(a, 1), "'", ""), ",", ".") & "',"
    StrSql = StrSql & Replace(grid.TextMatrix(a, 2), ",", ".") & ",'"
    StrSql = StrSql & grid.TextMatrix(a, 3) & "','"
    StrSql = StrSql & grid.TextMatrix(a, 4) & "')"
    Dbbase.Execute StrSql ', afetados
    '==> Marca Romaneio como feito
    NumerosPedidos = Split(grid.TextMatrix(a, 0), "-")
    For C = 0 To UBound(NumerosPedidos)
      StrSql = "Update alid050 set Romaneiogerado=1 where NumNf='" & NumerosPedidos(C) & "'"
     ExecutaSql StrSql
    Next
Next
GravaRomneio = CStr(Numero)

End Function

Function Imprime(Numero As Long)
On Error Resume Next
Dim LcFormula, LcCriterio As String, LcTipoSaida As Integer
Dim RsEmpresa As Recordset, RsOpcao As Recordset
Dim LcEmpresa, LcEndereco, LcFone, LcVer, LcCap, LcVer1 As String

'Abertura do relatório de vendas
    
    
CryRelatorio.DataFiles(0) = GLBase
CryRelatorio.ReportFileName = App.Path & "\Romaneio.rpt"
LcFormula = "{Romaneio.Codigo}=" & Numero
CryRelatorio.CopiesToPrinter = 1

CryRelatorio.DiscardSavedData = True
CryRelatorio.WindowTop = 50
CryRelatorio.WindowWidth = 700
CryRelatorio.WindowLeft = 50
CryRelatorio.WindowHeight = 500
CryRelatorio.WindowTitle = "Romaneio"

'CryRelatorio.Formulas(0) = "Empresa='" & LcEmpresa & "'"
'CryRelatorio.Formulas(1) = "Endereco='" & LcEndereco & "'"
'CryRelatorio.Formulas(2) = "Fone='(31)3388-1015 - Fax :3388-2520'"
'CryRelatorio.Formulas(3) = "Celular='Insc. Estadual: 062.608783.0021'"
'CryRelatorio.Formulas(4) = "email='CNPJ: 25.682.162/0001-88'"
'CryRelatorio.Formulas(5) = "titulo='Produtos'"
 
LcTipoSaida = 0
Me.Caption = LcCap
CryRelatorio.SelectionFormula = LcFormula

CryRelatorio.Destination = LcTipoSaida
CryRelatorio.PrintReport

'RsOpcao.Close
RsEmpresa.Close
Dbbase.Close
Set RsOpcao = Nothing
Set RsEmpresa = Nothing
Set Dbbase = Nothing
If CryRelatorio.LastErrorNumber > 0 Then MsgBox CryRelatorio.LastErrorString

End Function

