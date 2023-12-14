VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form FrmRelComissao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Comissao"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8625
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   8625
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport CryRelatorio 
      Left            =   4560
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame4 
      Caption         =   "Apos Imprimir"
      Height          =   1335
      Left            =   6240
      TabIndex        =   21
      Top             =   1920
      Visible         =   0   'False
      Width           =   2295
      Begin VB.OptionButton marcar 
         Caption         =   "Não Marcar como Pago"
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   840
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton marcarpg 
         Caption         =   "Marcar como Pago"
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Situção"
      Height          =   1335
      Left            =   4320
      TabIndex        =   17
      Top             =   1920
      Width           =   1695
      Begin VB.OptionButton todos 
         Caption         =   "Todas"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton naoPago 
         Caption         =   "Não Pago"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton pago 
         Caption         =   "Pago"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo "
      Height          =   1335
      Left            =   2040
      TabIndex        =   16
      Top             =   1920
      Width           =   2175
      Begin VB.OptionButton sintetico 
         Caption         =   "Sintético"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton analitico 
         Caption         =   "Analítico"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.TextBox codigo 
      Height          =   405
      Left            =   2640
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox Comissao 
      Height          =   315
      ItemData        =   "FrmRelComissao.frx":0000
      Left            =   120
      List            =   "FrmRelComissao.frx":0002
      TabIndex        =   0
      Top             =   480
      Width           =   3855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar F10"
      Height          =   615
      Left            =   6720
      TabIndex        =   9
      Top             =   840
      Width           =   1215
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
      Left            =   4800
      TabIndex        =   7
      Text            =   "1"
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Saída"
      Height          =   1335
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   1815
      Begin VB.OptionButton impressora 
         Caption         =   "Impressora"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton Video 
         Caption         =   "Video"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin MSMask.MaskEdBox Datai 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1320
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
      Left            =   2040
      TabIndex        =   2
      Top             =   1320
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
   Begin VB.CommandButton Command1 
      Caption         =   "&Confirma F3"
      Height          =   615
      Left            =   6720
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin VB.Line Line2 
      X1              =   6120
      X2              =   8640
      Y1              =   1800
      Y2              =   1800
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
      Left            =   120
      TabIndex        =   14
      Top             =   960
      Width           =   1185
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
      Left            =   2040
      TabIndex        =   13
      Top             =   960
      Width           =   1080
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
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1035
   End
   Begin VB.Line Line1 
      X1              =   6120
      X2              =   6120
      Y1              =   -240
      Y2              =   1800
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
      Left            =   4800
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   750
   End
End
Attribute VB_Name = "FrmRelComissao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type TipoVend
      codigo As String
      Nome As String
End Type
Private LcTamanho, a As Integer
Private MtVendedor() As TipoVend
Private RsComissao As Recordset, RsSintetico As Recordset
Private Rel       As New CrysRelComissao
Private RelA      As New CrysComissaoAnalitico
Function CarregaVendedor()
On Error GoTo errc
Dim RsVendedor As Recordset
AbreBase
Set RsVendedor = Dbbase.OpenRecordset("ALID200", dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcTamanho = 0
Do Until RsVendedor.EOF
   If Not IsNull(RsVendedor!Nome) Then
      ReDim Preserve MtVendedor(LcTamanho)
      MtVendedor(LcTamanho).codigo = RsVendedor!codigo
      MtVendedor(LcTamanho).Nome = RsVendedor!Nome
      Comissao.AddItem RsVendedor!Nome
      
      LcTamanho = LcTamanho + 1
   End If
   RsVendedor.MoveNext
Loop
If LcTamanho > 0 Then LcTamanho = LcTamanho - 1
'Comissao.AddItem "TODOS"
'Comissao.Text = "TODOS"
RsVendedor.Close
Set RsVendedor = Nothing
Exit Function
errc:
MsgBox err.Description & err.Number
'Resume 0
Exit Function
End Function
Function RecuperaNomeVendedor(codigo As String) As String
On Error GoTo errc
Dim RsVendedor As Recordset
AbreBase
Set RsVendedor = Dbbase.OpenRecordset("Select * from ALID200 where Codigo='" & codigo & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Dim LcNome As String
If Not RsVendedor.EOF Then
   If Not IsNull(RsVendedor!Nome) Then
      LcNome = RsVendedor!Nome
    End If
End If
RsVendedor.Close
Set RsVendedor = Nothing

RecuperaNomeVendedor = LcNome

Exit Function
errc:
MsgBox err.Description & err.Number
End Function
Private Sub analitico_Click()
'naoPago.Enabled = True
'pago.Enabled = True
'todos.Enabled = True

End Sub

Private Sub analitico_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Codigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Comissao_Click()
For a = 0 To LcTamanho
    If MtVendedor(a).Nome = Comissao.Text Then
       codigo.Text = MtVendedor(a).codigo
       Exit For
    End If
Next
If Len(Comissao.Text) = 0 Then Comissao.Text = ""

End Sub

Private Sub Comissao_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub
Sub IncluiDadosAlid201()
Dim StrSql      As String
Dim StrWhere    As String
On Error GoTo errGeraComissaoNova1

Dim RsNota As ADODB.Recordset
Dim RsSaida As ADODB.Recordset
Dim RsVendedor As Recordset
Dim rsCliente As Recordset

AbreBase

'==> Busca os dados da nfe

StrSql = "SELECT * from alid050"
StrSql = StrSql & "  WHERE  TipoOperacao like '1%' and (alid050.DTEMIS >= #" & Format(Datai.Text, "mm/dd/yy") & "# And alid050.DTEMIS <=#" & Format(Dataf.Text, "mm/dd/yy") & "#) AND ((alid050.status='Autorizado o uso da NF-e')or(alid050.status='Somente Gravada no Sistema')or (alid050.status='Em Lançamento'))"
If Len(codigo.Text) > 0 Then
   StrSql = StrSql & " and (vendedor='" & CLng(codigo.Text) & "')"
End If
'StrSql = StrSql & ")"
'Debug.Print StrSql
'MsgBox DEscricaoErro
'Somente Gravada no Sistema
Set RsNota = AbreRecordset(StrSql, True)

Do Until RsNota.EOF
  '==> Busca o cliente
  Dim LcNomeCliente As String
  Dim LcNomeVendedor As String
  Dim LcValorTotal As Single
  Dim LcItemBaixo As Boolean
  LcItemBaixo = False
  Dim LcComissao As Single
  Dim LcPercentual As Single
  
  LcNomeCliente = ""
  'StrSql = "Select * from alid001 where codigo='" & Right("00000" & RsNota!Cliente, 5) & "'"
  'Set rsCliente = Dbbase.OpenRecordset(StrSql)
  'If Not rsCliente.EOF Then
  '    LcNomeCliente = rsCliente!razaosoc & ""
  'End If
  'rsCliente.Close
  
  StrSql = "Select * from alid200 where codigo='" & Right("00000" & RsNota!vendedor, 5) & "'"
  Set RsVendedor = Dbbase.OpenRecordset(StrSql)
  If Not RsVendedor.EOF Then
      LcNomeVendedor = RsVendedor!Nome & ""
  End If
  RsVendedor.Close
  LcValorTotal = RsNota!ValorNota 'RsNota!Qtde * RsNota!VALUNIT
  LcPercentual = 1.5
  LcComissao = LcValorTotal * (LcPercentual / 100)
  
  '==> Insere a informacao no alid201
  StrSql = "Insert into alid201 (VENDEDOR,nf,produto,quantidade,VALORUNIT,VALORTOTAL,ITEMBAIXO"
  StrSql = StrSql & ",COMISSAO,DATAVENDA,CLIENTE,pago,percentual,valorpago,saldo,NomeCliente,NomeVendedor"
  StrSql = StrSql & ",DescontUnitario,com,NumeroNFE) values("
  StrSql = StrSql & "'" & Right("00000" & RsNota!vendedor, 5) & "',"
  StrSql = StrSql & "'" & RsNota!NumNf & "',"
  StrSql = StrSql & "0,"
  StrSql = StrSql & "0,"
  StrSql = StrSql & "0,"
  StrSql = StrSql & Replace(LcValorTotal, ",", ".") & ","
  StrSql = StrSql & LcItemBaixo & ","
  StrSql = StrSql & Replace(LcComissao, ",", ".") & ","
  StrSql = StrSql & "#" & Format(RsNota!DTEMIS, "mm/dd/yy") & "#,"
  StrSql = StrSql & "'" & Right("00000" & RsNota!Cliente, 5) & "',"
  StrSql = StrSql & "0,"
  StrSql = StrSql & Replace(LcPercentual, ",", ".") & ","
  StrSql = StrSql & "0,"
  StrSql = StrSql & "0,"
  StrSql = StrSql & "'" & Replace(RsNota!NomeCliente, "'", "''") & "',"
  StrSql = StrSql & "'" & Replace(LcNomeVendedor, "'", "''") & "',"
  StrSql = StrSql & "0,"
  StrSql = StrSql & "0,"
  StrSql = StrSql & "'" & RsNota!NumNf & "')"
  afetados = ExecutaSql(StrSql)
 
  RsNota.MoveNext
Loop
RsNota.Close
Set RsNota = Nothing
'===> puxa al's
StrSql = "SELECT * from saidas"
StrSql = StrSql & "  WHERE (saidas.DTEMIS >= #" & Format(Datai.Text, "mm/dd/yy") & "# And saidas.DTEMIS <=#" & Format(Dataf.Text, "mm/dd/yy") & "#) AND ((saidas.status='Autorizado o uso da NF-e')or(saidas.status='Somente Gravada no Sistema')or (saidas.status='EMITIDA'))"
If Len(codigo.Text) > 0 Then
   StrSql = StrSql & " and (vendedor='" & Right("00000" & codigo.Text, 5) & "')"
End If
'StrSql = StrSql & ")"
'Debug.Print StrSql
'MsgBox DEscricaoErro
'Somente Gravada no Sistema
Set RsNota = AbreRecordset(StrSql, True)

Do Until RsNota.EOF
  '==> Busca o cliente
  'Dim LcNomeCliente As String
  'Dim LcNomeVendedor As String
  'Dim LcValorTotal As Single
  'Dim LcItemBaixo As Boolean
  LcItemBaixo = False
  'Dim LcComissao As Single
  'Dim LcPercentual As Single
  
  LcNomeCliente = ""
  StrSql = "Select * from alid001 where codigo='" & Right("00000" & RsNota!Cliente, 5) & "'"
  Set rsCliente = Dbbase.OpenRecordset(StrSql)
  If Not rsCliente.EOF Then
      LcNomeCliente = rsCliente!RazaoSoc & ""
  End If
  rsCliente.Close
  
  StrSql = "Select * from alid200 where codigo='" & Right("00000" & RsNota!vendedor, 5) & "'"
  Set RsVendedor = Dbbase.OpenRecordset(StrSql)
  If Not RsVendedor.EOF Then
      LcNomeVendedor = RsVendedor!Nome & ""
  End If
  RsVendedor.Close
  LcValorTotal = RsNota!ValorNota 'RsNota!Qtde * RsNota!VALUNIT
  LcPercentual = 1.5
  LcComissao = LcValorTotal * (LcPercentual / 100)
  
  '==> Insere a informacao no alid201
  StrSql = "Insert into alid201 (VENDEDOR,nf,produto,quantidade,VALORUNIT,VALORTOTAL,ITEMBAIXO"
  StrSql = StrSql & ",COMISSAO,DATAVENDA,CLIENTE,pago,percentual,valorpago,saldo,NomeCliente,NomeVendedor"
  StrSql = StrSql & ",DescontUnitario,com,NumeroNFE) values("
  StrSql = StrSql & "'" & Right("00000" & RsNota!vendedor, 5) & "',"
  StrSql = StrSql & "'" & RsNota!NumNf & "',"
  StrSql = StrSql & "0,"
  StrSql = StrSql & "0,"
  StrSql = StrSql & "0,"
  StrSql = StrSql & Replace(LcValorTotal, ",", ".") & ","
  StrSql = StrSql & LcItemBaixo & ","
  StrSql = StrSql & Replace(LcComissao, ",", ".") & ","
  StrSql = StrSql & "#" & Format(RsNota!DTEMIS, "mm/dd/yy") & "#,"
  StrSql = StrSql & "'" & Right("00000" & RsNota!Cliente, 5) & "',"
  StrSql = StrSql & "0,"
  StrSql = StrSql & Replace(LcPercentual, ",", ".") & ","
  StrSql = StrSql & "0,"
  StrSql = StrSql & "0,"
  StrSql = StrSql & "'" & Replace(LcNomeCliente, "'", "''") & "',"
  StrSql = StrSql & "'" & Replace(LcNomeVendedor, "'", "''") & "',"
  StrSql = StrSql & "0,"
  StrSql = StrSql & "0,"
  StrSql = StrSql & "'" & RsNota!NumNf & "')"
  afetados = ExecutaSql(StrSql)
 'MsgBox DEscricaoErro
  RsNota.MoveNext
Loop

Exit Sub
errGeraComissaoNova1:
Stop
MsgBox err.Description & " " & err.Number
Resume 0
End Sub
Sub GeraComissaoNova()
Dim StrSql      As String
Dim StrWhere    As String
On Error GoTo errGeraComissaoNova

Dim Rs As ADODB.Recordset

StrSql = "Select * from alid201"
StrSql = StrSql & " where datavenda between '" & Format(Datai.Text, "yyyy-mm-dd") & "' and '" & Format(Dataf.Text, "yyyy-mm-dd") & "'"
If Len(codigo.Text) > 0 Then
   StrSql = StrSql & " And VENDEDOR='" & codigo.Text & "'"
End If

StrSql = StrSql & " order by alid201.codigo"
Debug.Print StrSql
LcCap = Me.Caption
Me.Caption = "Aguarde, processando os dados..."
Screen.MousePointer = 11
Screen.MousePointer = 0
Me.Caption = LcCap
Screen.MousePointer = vbHourglass
'==>Exclui os itens
Dim StrExclui As String
'StrExclui = "Delete from alid201"

Set Rs = AbreRecordset(StrSql, True)
Debug.Print StrSql
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
Exit Sub
errGeraComissaoNova:
MsgBox err.Description & err.Number
Resume Next
End Sub
    Private Function CalculaPercentualComissao(codigoproduto As Integer, ValorUnitarioVenda As Currency, QuantidadeUnidade As Currency, ByRef ValorMinimo As Currency, ByRef LcLimiteVenda As Currency) As Currency
        On Error Resume Next
        'Dim RsProduto As ADODB.Recordset
       
        Dim StrSql As String
        Dim LcPRecoAntigo As Currency
        Dim MinimoVenda As Currency
        Dim QuantUnidade As Currency
        Dim Percentual As Currency
        Dim LimiteVenda As Currency
        '==> abre a tabela produtos para saber o minimo
        StrSql = "Select * from PRodutos where codigo=" & codigoproduto
        Dim RsProduto As ADODB.Recordset
        Set RsProduto = AbreRecordset(StrSql, True)
        
        If Not RsProduto.EOF Then
            LcPRecoAntigo = RsProduto("Preco") 'RsProduto!Preco
            MinimoVenda = RsProduto("MinimoVenda") 'RsProduto!MinimoVenda
            QuantUnidade = RsProduto("QtdMedida") 'RsProduto!QtdMedida
          If IsNumeric(RsProduto("LimiteVenda")) Then LimiteVenda = RsProduto("LimiteVenda")
        
        End If
        ValorMinimo = MinimoVenda
        LcLimiteVenda = LimiteVenda
        ' Dr.Close()
        '==>Calcula o Valor unitario da unidade
        If QuantUnidade <> QuantidadeUnidade Then
            '==> Te que calcular os precos
            MinimoVenda = (MinimoVenda / QuantUnidade)
            ValorUnitarioVenda = ValorUnitarioVenda / QuantidadeUnidade
            LcPRecoAntigo = LcPRecoAntigo / QuantidadeUnidade
            LimiteVenda = LimiteVenda / QuantidadeUnidade
        End If
        Percentual = 1
        'If ValorUnitarioVenda < LimiteVenda Then
        '    '===> Comissao é 0,5 %
        '    Percentual = 0.5
        'ElseIf (Math.Round(ValorUnitarioVenda, 4) >= Math.Round(LimiteVenda, 4)) And (Math.Round(ValorUnitarioVenda, 4) < Math.Round(MinimoVenda, 4)) Then 'ValorUnitarioVenda < MinimoVenda Then
        '    Percentual = 1
       ' ElseIf (Math.Round(ValorUnitarioVenda, 4) >= Math.Round(MinimoVenda, 4)) And (Math.Round(ValorUnitarioVenda, 4) < Math.Round(LcPRecoAntigo, 4)) Then
        '    '===> Comissao é 1 %
        '    Percentual = 1.5
        'ElseIf (Math.Round(ValorUnitarioVenda, 4) >= Math.Round(LcPRecoAntigo, 4)) Then
        '    '===> Comissao é 1,5 %
        '    Percentual = 2
        'End If
        'If ValorUnitarioVenda < MinimoVenda Then
            '===> Comissao é 0,5 %
          '  Percentual = 0.5
        'ElseIf (Math.Round(ValorUnitarioVenda, 4) >= Math.Round(MinimoVenda, 4)) And (Math.Round(ValorUnitarioVenda, 4) < Math.Round(LcPRecoAntigo, 4)) Then
            '===> Comissao é 1 %
            'Percentual = 1
        'ElseIf (Math.Round(ValorUnitarioVenda, 4) >= Math.Round(LcPRecoAntigo, 4)) Then
            '===> Comissao é 1,5 %
           ' Percentual = 1.5
       ' End If
        
        CalculaPercentualComissao = Percentual
    End Function
Sub GravaComissao(NF As String)
  Dim a As Integer
  Dim PercentualDesconto As Single
  Dim LcValorComissao As Single
  ' Dim PercentualComissao As String
  Dim Resposta As Boolean
  Dim RsNota As ADODB.Recordset
  Dim RsdadosNota As ADODB.Recordset
  Set RsNota = AbreRecordset("Select * From alid050 where NUMNF='" & NF & "'", True)
  Set RsdadosNota = AbreRecordset("Select * From alid052 where NUMNF='" & NF & "'", True)
        '==> Inicia a Base de dados
 '==> Determina se tem item baixo na nota
 If Not RsNota.EOF And Not RsdadosNota.EOF Then
   Dim LcItemBaixo As Boolean '= dadosNota.ItemBaixo
        GlErro = ""
       
            '==> Calcula po Percentual de desconto na nota
            If RsNota("Desconto") > 0 Then
                PercentualDesconto = CSng(RsNota("Desconto")) / CSng(RsNota("ValorNota"))
            End If
            '=> Efetua a Exclusão dos dados desta nota na comissao.
            'StrSql = "DELETE from alid201 where nf='" & dadosNota.numeronota & "'"
            
            Do Until RsdadosNota.EOF
                Dim LcValorMinimo As Currency
                Dim LcPercentual As Currency
                Dim LcLimiteVenda As Currency
                LcPercentual = CalculaPercentualComissao(RsdadosNota("codProd"), RsdadosNota("VALUNIT"), RsdadosNota("QTDUM"), LcValorMinimo, LcLimiteVenda)
                StrSql = "Insert into alid201 (VENDEDOR,nf,produto,quantidade,com,VALORUNIT,VALORTOTAL,ITEMBAIXO,"
                StrSql = StrSql & " COMISSAO,DATAVENDA,CLIENTE,percentual,valorpago,NomeCliente,NomeVendedor,DescontUnitario,valorMinimo,NumeroNFE,LimiteVenda) Values ("
                StrSql = StrSql & "'" & Right("00000" & RsNota("Vendedor"), 5) & "',"
                StrSql = StrSql & "'" & NF & "',"
                StrSql = StrSql & RsdadosNota("codProd") & ","
                StrSql = StrSql & Replace(RsdadosNota("QTDE"), ",", ".") & ","
                StrSql = StrSql & RsdadosNota("QTDUM") & ","
                StrSql = StrSql & Replace(Math.Round(RsdadosNota("VALUNIT"), 2), ",", ".") & ","
                StrSql = StrSql & Replace(Math.Round(RsdadosNota("VALUNIT") * RsdadosNota("QTDE"), 2), ",", ".") & ","
                StrSql = StrSql & LcItemBaixo & ","
                LcValorComissao = Math.Round((RsdadosNota("VALUNIT") * RsdadosNota("QTDE")) * (LcPercentual / 100), 2)
                StrSql = StrSql & Replace(LcValorComissao, ",", ".") & ","
                StrSql = StrSql & "'" & Format(CDate(RsNota("DTEMIS")), "yyyy-mm-dd") & "',"
                StrSql = StrSql & CLng(RsNota("Cliente")) & ","
                StrSql = StrSql & Replace(LcPercentual, ",", ".") & ","
                StrSql = StrSql & "0,"
                StrSql = StrSql & "'" & Replace(RsNota("NomeCliente"), "'", "''") & "',"
                StrSql = StrSql & "'" & Replace(RecuperaNomeVendedor(Right("00000" & RsNota("Vendedor"), 5)), "'", "''") & "',"
                StrSql = StrSql & Replace(CStr(PercentualDesconto * RsdadosNota("VALUNIT")), ",", ".") & ","
                StrSql = StrSql & Replace(LcValorMinimo, ",", ".") & ","
                StrSql = StrSql & "'" & NF & "',"
                StrSql = StrSql & Replace(LcLimiteVenda, ",", ".") & ")"
                Dim afetados As Integer
                'Debug.Print StrSql
                conexaoAdo.Execute StrSql
                
                RsdadosNota.MoveNext
            Loop
   End If
 End Sub
Sub verificaComissaoNaoLancada()
Dim StrSql      As String
Dim StrWhere    As String
Dim LcTotal As Long
Dim a As Long
Me.Caption = "Verificando Lançamento das Notas Fiscais..."
DoEvents

Dim Rs As ADODB.Recordset
StrSql = "SELECT alid050.NUMNF, alid050.DTEMIS, alid050.status, alid201.NF"
StrSql = StrSql & " FROM alid050 LEFT JOIN alid201 ON alid050.NUMNF = alid201.NF"
StrSql = StrSql & " WHERE (((alid050.DTEMIS) Between '" & Format(Datai.Text, "yyyy-mm-dd") & "' and '" & Format(Dataf.Text, "yyyy-mm-dd") & "') AND ((alid050.status)='Autorizado o uso da NF-e') AND ((alid201.NF) Is Null) and TipoOperacao like '1%');"
Debug.Print StrSql
Set Rs = AbreRecordset(StrSql, True)
LcTotal = Rs.RecordCount
Do Until Rs.EOF
    a = a + 1
    Me.Caption = "Verificando lançamento da comissão NF:" & Rs!NumNf & " Registro " & a & " de " & LcTotal
    DoEvents
    GravaComissao Rs("NumNF")
   Rs.MoveNext
Loop

End Sub
Sub ExcluiEntradas()
Dim StrSql      As String
Dim StrWhere    As String
Dim LcTotal As Long
Dim a As Long
Me.Caption = "Verificando Lançamento das Notas Fiscais..."
DoEvents

Dim Rs As ADODB.Recordset
StrSql = "SELECT alid050.NUMNF, alid050.DTEMIS, alid050.status, alid201.NF"
StrSql = StrSql & " FROM alid050 LEFT JOIN alid201 ON alid050.NUMNF = alid201.NF"
StrSql = StrSql & " WHERE (((alid050.DTEMIS) Between '" & Format(Datai.Text, "yyyy-mm-dd") & "' and '" & Format(Dataf.Text, "yyyy-mm-dd") & "') AND ((alid050.status)='Autorizado o uso da NF-e') and TipoOperacao like '0%');"
Debug.Print StrSql
Set Rs = AbreRecordset(StrSql, True)
LcTotal = Rs.RecordCount
Do Until Rs.EOF
    a = a + 1
    Me.Caption = "Verificando lançamento da comissão NF:" & Rs!NumNf & " Registro " & a & " de " & LcTotal
    DoEvents
    StrSql = "Delete from alid201 where NF='" & Rs!NF & "'"
    ExecutaSql StrSql
   Rs.MoveNext
Loop

End Sub
Sub AcertaComissao()
Dim StrSql      As String
Dim StrWhere    As String
Dim Rs As ADODB.Recordset
Dim Rsp As ADODB.Recordset
Dim LcTotal As Long
Dim a As Long
Dim LcCap As String
LcCap = Me.Caption
StrSql = "SELECT * from alid201 where DATAVENDA between '" & Format(Datai.Text, "yyyy-mm-dd") & "' and '" & Format(Dataf.Text, "yyyy-mm-dd") & "'"
Set Rs = AbreRecordset(StrSql, True)
LcTotal = Rs.RecordCount
Do Until Rs.EOF
   a = a + 1
   '===> Busca as informações no cadastro de produto
   'If Rs!Produto = 6046 Then Stop
   Me.Caption = "Verificando a comissão da nota :" & Rs!NF & " Registro:" & a & " de " & LcTotal
   DoEvents
   StrSql = "Select * from Produtos where codigo=" & Rs!produto
   Dim ValorUnitarioVenda As Double
   Dim MinimoVenda As Currency
   Dim QuantUnidPadrao As Currency
   Dim PrecoBase As Currency
   Dim QuantUnidVenda As Currency
   Dim LimiteVenda As Currency
   Dim Perc As Currency
   Dim LcValorComissao As Currency
  ' Dim RsProduto As ADODB.Recordset
   'Set RsProduto = AbreRecordset(StrSql, True)
   
   ValorUnitarioVenda = Math.Round(Rs!ValorUnit, 2)
   QuantUnidVenda = Math.Round(Rs!Com, 2)
   MinimoVenda = Math.Round(Rs!ValorMinimo, 2)
   LimiteVenda = Math.Round(Rs!LimiteVenda, 2)
   
   '==. Recupera as Informações do Produto
   Set Rsp = AbreRecordset(StrSql, True)
   If Not Rsp.EOF Then
      QuantUnidPadrao = Math.Round(Rsp!QtdMedida, 2)
      PrecoBase = Math.Round(Rsp!Preco, 2)
      If IsNumeric(Rsp!MinimoVenda) Then If MinimoVenda <= 0 Then MinimoVenda = Rsp!MinimoVenda
      If IsNumeric(Rsp!LimiteVenda) Then If LimiteVenda <= 0 Then LimiteVenda = Rsp!LimiteVenda
   Else
      QuantUnidPadrao = 0
      PrecoBase = 0
   End If
   '---> Veririfica se foi vendido na quantidade padrao
   If QuantUnidVenda <> QuantUnidPadrao Then
        If QuantUnidPadrao > 0 Then
          MinimoVenda = Math.Round(MinimoVenda / QuantUnidPadrao, 2)
          PrecoBase = Math.Round(PrecoBase / QuantUnidPadrao, 2)
          LimiteVenda = Math.Round(LimiteVenda / QuantUnidPadrao, 2)
        End If
      
   End If
   
    'If ValorUnitarioVenda < LimiteVenda Then
    '        '===> Comissao é 0,5 %
    '        Percentual = 0.5
    '    ElseIf (Math.Round(ValorUnitarioVenda, 4) >= Math.Round(LimiteVenda, 4)) And (Math.Round(ValorUnitarioVenda, 4) < Math.Round(MinimoVenda, 4)) Then 'ValorUnitarioVenda < MinimoVenda Then
    '        Percentual = 1
    '    ElseIf (Math.Round(ValorUnitarioVenda, 4) >= Math.Round(MinimoVenda, 4)) And (Math.Round(ValorUnitarioVenda, 4) < Math.Round(PrecoBase, 4)) Then
    '        '===> Comissao é 1 %
    '        Percentual = 1.5
    '    ElseIf (Math.Round(ValorUnitarioVenda, 4) >= Math.Round(PrecoBase, 4)) Then
    '       '===> Comissao é 1,5 %
    '        Percentual = 2
    '    End If
   Percentual = 1
  
   '===> Recalcula a comissao
   LcValorComissao = Math.Round(Rs!ValorTotal * (Percentual / 100), 2)
   '===> Altera o banco
   StrSql = "Update alid201 set COMISSAO=" & Replace(CStr(LcValorComissao), ",", ".") & ",percentual=" & Replace(CStr(Percentual), ",", ".") & ",LimiteVenda=" & Replace(CStr(LimiteVenda), ",", ".") & " where codigo=" & Rs!codigo
             'Update alid201 set COMISSAO=5.36,percentual=0.5 where codigo=3234162
   'Debug.Print StrSql
   Dim afetados As Integer
    conexaoAdo.Execute StrSql, afetados
    DoEvents
  Rs.MoveNext
Loop
Me.Caption = LcCap


End Sub
Sub RelancaComissao()
On Erro GoTo errOcorrido
Dim a As Integer
  Dim PercentualDesconto As Currency
  Dim LcValorComissao As Currency
  ' Dim PercentualComissao As String
  Dim Resposta As Boolean
  Dim RsNota As ADODB.Recordset
  Dim RsdadosNota As ADODB.Recordset
  Dim StrSql As String
  Dim StrExclui As String
  Dim afetados As Long
  StrExclui = "delete from alid201"
  conexaoAdo.Execute StrExclui, afetados
  
  StrSql = "select codigo,NUMNF,DTEmis,Natureza,valorproduto,valornota,status,Desconto,finalidadeEmissao,FomaEmissao,tipoOperacao,vendedor,cliente,nomecliente from alid050 where alid050.DTEMIS Between '" & Format(Datai.Text, "yyyy-mm-dd") & "' and '" & Format(Dataf.Text, "yyyy-mm-dd") & "' and Status<>'CANCELADA' and Status<>'INUTILIZADA'"
  StrSql = StrSql & " and finalidadeEmissao='1- NF-e normal' and tipoOperacao='1 - Saida'and (not (Status like'%Denegad%'))"
  If IsNumeric(codigo.Text) Then
     If CInt(codigo.Text) > 0 Then
         StrSql = StrSql & "  and (vendedor=" & codigo.Text & ")"
     End If
  End If
  
  Set RsNota = AbreRecordset(StrSql, True)
 'Debug.Print StrSql
        '==> Inicia a Base de dados
 '==> Determina se tem item baixo na nota

Do Until RsNota.EOF
   Dim LcItemBaixo As Boolean '= dadosNota.ItemBaixo
        GlErro = ""
            PercentualDesconto = 0
            Set RsdadosNota = AbreRecordset("Select * From alid052 where NUMNF='" & RsNota("NUMNF") & "' order by item", True)
            '==> Calcula po Percentual de desconto na nota
            If RsNota("Desconto") > 0 Then
                PercentualDesconto = Math.Round(CCur(RsNota("Desconto")) / CCur(RsNota("valorproduto")), 4)
            End If
            '==> Quantidade de registros filhos
            Dim TotalReg As Long
            Dim ValorDesconto As Currency
            Dim diferencaDesc As Currency
           
            TotalReg = RsdadosNota.RecordCount
            
            Dim regs As Long
            regs = 0
            ValorDesconto = 0
            diferencaDesc = 0
            LcValorTotal = 0
            '=> Efetua a Exclusão dos dados desta nota na comissao.
            'StrSql = "DELETE from alid201 where nf='" & dadosNota.numeronota & "'"
            Me.Caption = "Processando nfe:" & RsNota("NUMNF")
            Do Until RsdadosNota.EOF
            regs = regs + 1
              If RsNota("Desconto") > 0 Then
                 If TotalReg = regs Then
                   diferencaDesc = Math.Round(RsNota("Desconto"), 2) - Math.Round(ValorDesconto, 2)
                 End If
              End If
               
               DoEvents
                Dim LcValorMinimo As Currency
                Dim LcPercentual As Currency
                Dim LcLimiteVenda As Currency
                Dim LcDesconto As Currency
                'Dim LcValorTotal As Double
                LcDesconto = 0
                 LcDesconto = Math.Round(PercentualDesconto * (RsdadosNota("VALUNIT") * RsdadosNota("QTDE")), 2)
                 ValorDesconto = Math.Round(ValorDesconto + LcDesconto, 2)
                If diferencaDesc > 0 Then
                   LcDesconto = Math.Round(LcDesconto, 2)
                Else
                   LcDesconto = Math.Round(LcDesconto, 2)
                End If
                LcValorTotal = Math.Round((RsdadosNota("VALUNIT") * RsdadosNota("QTDE")) - LcDesconto, 2)
                LcPercentual = CalculaPercentualComissao(RsdadosNota("codProd"), RsdadosNota("VALUNIT"), RsdadosNota("QTDUM"), LcValorMinimo, LcLimiteVenda)
                StrSql = "Insert into alid201 (VENDEDOR,nf,produto,quantidade,com,VALORUNIT,VALORTOTAL,ITEMBAIXO,"
                StrSql = StrSql & " COMISSAO,DATAVENDA,CLIENTE,percentual,valorpago,NomeCliente,NomeVendedor,DescontUnitario,valorMinimo,NumeroNFE,LimiteVenda) Values ("
                StrSql = StrSql & "'" & Right("00000" & RsNota("Vendedor"), 5) & "',"
                StrSql = StrSql & "'" & RsNota("NUMNF") & "',"
                StrSql = StrSql & RsdadosNota("codProd") & ","
                StrSql = StrSql & Replace(RsdadosNota("QTDE"), ",", ".") & ","
                StrSql = StrSql & RsdadosNota("QTDUM") & ","
                StrSql = StrSql & Replace(Math.Round(RsdadosNota("VALUNIT"), 2), ",", ".") & ","
                StrSql = StrSql & Replace(LcValorTotal, ",", ".") & ","
                StrSql = StrSql & LcItemBaixo & ","
                LcValorComissao = Math.Round(LcValorTotal * (LcPercentual / 100), 2)
                StrSql = StrSql & Replace(LcValorComissao, ",", ".") & ","
                StrSql = StrSql & "'" & Format(CDate(RsNota("DTEMIS")), "yyyy-mm-dd") & "',"
                StrSql = StrSql & CLng(RsNota("Cliente")) & ","
                StrSql = StrSql & Replace(LcPercentual, ",", ".") & ","
                StrSql = StrSql & "0,"
                StrSql = StrSql & "'" & Replace(RsNota("NomeCliente"), "'", "''") & "',"
                StrSql = StrSql & "'" & Replace(RecuperaNomeVendedor(Right("00000" & RsNota("Vendedor"), 5)), "'", "''") & "',"
                StrSql = StrSql & Replace(CStr(PercentualDesconto * RsdadosNota("VALUNIT")), ",", ".") & ","
                StrSql = StrSql & Replace(LcValorMinimo, ",", ".") & ","
                StrSql = StrSql & "'" & RsNota("NUMNF") & "',"
                StrSql = StrSql & Replace(LcLimiteVenda, ",", ".") & ")"
                
                'Debug.Print StrSql
                conexaoAdo.Execute StrSql
                
                RsdadosNota.MoveNext
            Loop
            DoEvents
            RsNota.MoveNext
   Loop
   Exit Sub
errOcorrido:
   MsgBox err.Description
   Resume 0
End Sub
Sub GeraComissaoNovaAnalitico()
Dim StrSql      As String
Dim StrWhere    As String
On Error GoTo errGeraComissaoNova
Dim Rs As ADODB.Recordset
StrSql = "SELECT alid201.VENDEDOR, alid201.NF, alid201.PRODUTO, alid201.QUANTIDADE, alid201.VALORUNIT, alid201.LimiteVenda as LIMITEVENDA, alid201.VALORTOTAL, alid201.ITEMBAIXO, alid201.COMISSAO, alid201.DATAVENDA, alid201.CLIENTE, alid201.codigo, alid201.pago, alid201.Fornecedor, alid201.percentual, alid201.valorpago, alid201.saldo, alid201.NomeCliente, alid201.NomeVendedor, alid201.DescontUnitario, alid201.com, alid201.NumeroNFE, alid201.valorMinimo, produtos.NOME as nome"
StrSql = StrSql & " FROM alid201 INNER JOIN produtos ON alid201.PRODUTO = produtos.codigo "
StrSql = StrSql & " where datavenda between '" & Format(Datai.Text, "yyyy-mm-dd") & "' and '" & Format(Dataf.Text, "yyyy-mm-dd") & "'"
If Len(codigo.Text) > 0 Then
   StrSql = StrSql & " And VENDEDOR='" & codigo.Text & "'"
End If
StrSql = StrSql & " order by alid201.codigo"
Set Rs = AbreRecordset(StrSql, True)
Debug.Print StrSql
LcCap = Me.Caption
'==> Acerta o nome do vendedor
Me.Caption = "Aguarde, processando os dados..."
Screen.MousePointer = 11
Screen.MousePointer = 0
Me.Caption = LcCap
Screen.MousePointer = vbHourglass
'==>Exclui os itens
'Dim StrExclui As String
'StrExclui = "Delete from alid201"


'AcessoAdo.ExecutaSql StrExclui
'incluiDadosAlid201

Debug.Print StrSql
Load Relatorios
With Relatorios
     RelA.DiscardSavedData
     RelA.Database.SetDataSource Rs
     .CRViewer1.ReportSource = RelA
     setaformulaA
      .CRViewer1.ViewReport
End With
Relatorios.Show

Screen.MousePointer = vbDefault
Exit Sub
errGeraComissaoNova:
MsgBox err.Description & err.Number
Resume Next
End Sub
Sub setaformulaA()
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
   lctitulo = "Relatório de Comissões Período de : " & Datai.Text & " à " & Dataf.Text
   Else
   lctitulo = "Relatório de Comissões"
End If
If Len(codigo.Text) > 0 And Len(Comissao.Text) > 0 Then
   lctitulo = lctitulo & " Vend:" & CInt(codigo.Text) & " - " & Comissao.Text
ElseIf Len(Comissao.Text) > 0 Then
   lctitulo = lctitulo & " Vend:" & Comissao.Text
End If
With RelA
'Exit Sub
For a = 1 To .FormulaFields.Count
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("Fone") Then .FormulaFields(a).Text = "totext('" & LcFone & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("EMPRESA") Then .FormulaFields(a).Text = "totext('" & LcEmpresa & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("ENDERECO") Then .FormulaFields(a).Text = "totext('" & LcEndereco & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("Titulo") Then
           .FormulaFields(a).Text = "totext('" & lctitulo & "')"
        End If
    Next
End With
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
   lctitulo = "Relatório de Comissões Período de : " & Datai.Text & " à " & Dataf.Text
Else
   lctitulo = "Relatório de Comissões"
End If
If Len(codigo.Text) > 0 And Len(Comissao.Text) > 0 Then
   lctitulo = lctitulo & " Vend:" & CInt(codigo.Text) & " - " & Comissao.Text
ElseIf Len(Comissao.Text) > 0 Then
   lctitulo = lctitulo & " Vend:" & Comissao.Text
End If

With Rel
'Exit Sub
For a = 1 To .FormulaFields.Count
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("Fone") Then .FormulaFields(a).Text = "totext('" & LcFone & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("EMPRESA") Then .FormulaFields(a).Text = "totext('" & LcEmpresa & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("ENDERECO") Then .FormulaFields(a).Text = "totext('" & LcEndereco & "')"
        If UCase(.FormulaFields(a).FormulaFieldName) = UCase("Titulo") Then
           .FormulaFields(a).Text = "totext('" & lctitulo & "')"
        End If
    Next
End With
End Sub
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
Sub Carregadados()
On Error GoTo ErroCarrega

Dim RsNota As ADODB.Recordset
Dim RsNotaMdb As DAO.Recordset
AbreBase
'abreconexao
LcSql = "Select * from produtos"
Set RsNota = AbreRecordsetRel(LcSql, RsNota)
Set RsNotaMdb = Dbbase.OpenRecordset("Select * from produtos")
RsNota.Requery
'===> Apagando Registros antigos
Do Until RsNotaMdb.EOF
    RsNotaMdb.Delete
    RsNotaMdb.MoveNext
Loop
Do Until RsNota.EOF
    RsNotaMdb.AddNew
    For C = 0 To RsNota.Fields.Count - 1
        LcNome = RsNota.Fields(C).Name
        RsNotaMdb(LcNome) = RsNota.Fields(C)
        DoEvents
    Next
    RsNotaMdb.Update
    RsNota.MoveNext
    DoEvents
Loop
RsNota.Close
'FechaConexao
RsNotaMdb.Close

Exit Sub
ErroCarrega:
MsgBox err.Number & " " & err.Description
'Resume 0
End Sub
Sub GeraComissaoCrystalXI()
Dim LcWhere As String
Dim StrSql As String
Dim Rs As ADODB.Recordset

If Len(codigo.Text) = 0 Then codigo.Text = 0
LcWhere = " Where (status='Autorizado o uso da NF-e' or status='EMITIDA') "
If IsDate(Datai.Text) Then
   If IsDate(Dataf.Text) Then
       If Len(LcWhere) = 0 Then LcWhere = "Where " Else LcWhere = LcWhere & " and "
       'LcWhere += "(DATAVENDA Between #" & Format(DataI, "MM/dd/yy") & "# And #" & Format(dataf, "MM/dd/yy") & "#)"
       LcWhere = LcWhere & "(DTEMIS Between #" & Format(Datai.Text, "mm/dd/yy") & "# And #" & Format(Dataf.Text, "mm/dd/yy") & "#)"
   Else
       If Len(LcWhere) = 0 Then LcWhere = "Where " Else LcWhere = LcWhere & " and "
          'LcWhere += "(DATAVENDA = #" & Format(DataI, "MM/dd/yy") & "#)"
          LcWhere = LcWhere & "(DTEMIS = #" & Format(Datai.Text, "mm/dd/yy") & "#)"
       End If
   End If
   If CInt(codigo.Text) > 0 Then
      If Len(LcWhere) = 0 Then LcWhere = "Where " Else LcWhere = LcWhere & " and "
      LcWhere = "lcwhere & (VENDEDOR='" & codigo.Text & "')"
   End If
   StrSql = "SELECT alid050.Codigo, alid050.NUMNF, alid050.DTEMIS, alid050.CLIENTE, alid050.ValorNota, alid050.Vendedor, alid050.status, alid052.QTDE, alid052.codProd, alid052.descricao, alid052.VALUNIT "
   StrSql = StrSql & " FROM alid050 INNER JOIN alid052 ON alid050.codigo = alid052.CodigoNota"
   'WHERE (((alid050.DTEMIS) Between #2/1/2011# And #2/28/2011#) AND ((alid050.status)="Autorizado o uso da NF-e"));
   StrSql = StrSql & LcWhere & " order by alid050.DTEMIS"
      
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
Private Sub Command1_Click()
Dim LcFormula, LcCriterio As String, LcTipoSaida As Integer
Dim RsEmpresa As Recordset, RsBaixa As Recordset, RsOpcao As Recordset
Dim LcEmpresa, LcEndereco, LcFone, LcCelular, Lccelular1, Lcemail, LcVer, LcVer1, LcCap As String
LcCap = Me.Caption
Me.MousePointer = 11
Me.Caption = "Aguarde, processando a comissão..."
DoEvents
If Not IsDate(Datai.Text) Then
   MsgBox "A Data Inicial Não é Válida", 64, "Aviso"
   Exit Sub
End If
If Not IsDate(Dataf.Text) Then
   MsgBox "A Data Final Não é Válida", 64, "Aviso"
   Exit Sub
End If
RelancaComissao
'verificaComissaoNaoLancada
'ExcluiEntradas
AcertaComissao
   If analitico.Value Then
      GeraComissaoNovaAnalitico
   Else
      GeraComissaoNova
   End If
  Me.Caption = LcCap
  Me.MousePointer = 0
DoEvents
   Exit Sub
'End If
AbreBase

LcBaixa = "select * from Alid201 where VENDEDOR='" & codigo.Text & "' and DATAVENDA>=#" & Datai.Text & "# and DATAVENDA <=#" & Dataf.Text & "#"
Set RsEmpresa = Dbbase.OpenRecordset("Empresa", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsBaixa = Dbbase.OpenRecordset(LcBaixa)

'Set RsOpcao = Dbbase.OpenRecordset("Opcoes", dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcCap = Me.Caption
Me.Caption = "Aguarde, Gerando o Relatório..."
If Not RsEmpresa.EOF Then
   LcEmpresa = RsEmpresa!Razao
   LcEndereco = RsEmpresa!Endereco & " " & RsEmpresa!Bairro
   LcFone = RsEmpresa!Fone
End If
'If Not RsOpcao.EOF Then
 '  LcVer = RsOpcao!msg
 '  LcVer1 = RsOpcao!Msg1
'End If

    'Abertura do relatório de vendas
        
    CryRelatorio.DataFiles(0) = GLBase
    If analitico Then
       lctitulo = "Relatório de Comissões << ANALÍTICO >>"
       If GlImprimeSemLinha Then
          CryRelatorio.ReportFileName = App.Path & "\comissao.rpt"
       Else
          CryRelatorio.ReportFileName = App.Path & "\comissaosl.rpt"
       End If
       CryRelatorio.SortFields(0) = "+{ALID201.nf}"
    Else
       lctitulo = "Relatório de Comissões << SINTÉTICO >>"
       GeraSintetico
       If GlImprimeSemLinha Then
          CryRelatorio.ReportFileName = App.Path & "\comissaosintetico.rpt"
       Else
          CryRelatorio.ReportFileName = App.Path & "\comissaosinteticosl.rpt"
       End If
       CryRelatorio.SortFields(0) = ""
    End If
    'CryRelatorio.SortFields(0) = "+{ALID201.VENDEDORr}"
    
    CryRelatorio.CopiesToPrinter = Val(copias.Text)
    If analitico And Comissao.Text <> "TODOS" Then
       LcFormula = "{ALID201.VENDEDOR} = '" & codigo.Text & "'"
    End If

  '== Inicio Filtro
  strData = CDate(Format(Datai.Text, "dd/mm/yyyy"))
  LcAno = Year(strData)
  LcMes = Month(strData)
  LcDia = Day(strData)
  LcDataInicio = LcAno & "," & LcMes & "," & LcDia
  LcChav1 = " date(" & LcDataInicio & ")"
         
  strData = CDate(Format(Dataf.Text, "dd/mm/yyyy"))
  LcAno = Year(strData)
  LcMes = Month(strData)
  LcDia = Day(strData)
  LcDataInicio = LcAno & "," & LcMes & "," & LcDia
  LcChav2 = " date(" & LcDataInicio & ")"
  If Len(LcFormula) <> 0 Then LcFormula = LcFormula & " And "
  
  
  If analitico Then LcFormula = LcFormula & "{ALID201.DATAVENDA} >=" & LcChav1 & " And {ALID201.DATAVENDA} <=" & LcChav2
  
  If analitico Then
     If pago Then
        If Len(LcFormula) <> 0 Then LcFormula = LcFormula & " And "
        LcFormula = LcFormula & "{ALID201.pago}=True"
     End If
     If naoPago Then
        If Len(LcFormula) <> 0 Then LcFormula = LcFormula & " And "
        LcFormula = LcFormula & "{ALID201.pago}=False"
     End If
  Else
     If pago Then
        If Len(LcFormula) <> 0 Then LcFormula = LcFormula & " And "
        LcFormula = LcFormula & "{Sintetico.pago}=True"
     End If
     If naoPago Then
        If Len(LcFormula) <> 0 Then LcFormula = LcFormula & " And "
        LcFormula = LcFormula & "{Sintetico.pago}=False"
     End If
  
   End If
'== fim filtro
'== fim filtro
CryRelatorio.DiscardSavedData = True
CryRelatorio.WindowTop = 50
CryRelatorio.WindowWidth = 700
CryRelatorio.WindowLeft = 50
CryRelatorio.WindowHeight = 500
CryRelatorio.WindowState = crptMaximized
CryRelatorio.WindowTitle = lctitulo

 CryRelatorio.Formulas(0) = "Empresa='" & LcEmpresa & "'"
 CryRelatorio.Formulas(1) = "Endereco='" & LcEndereco & "'"
 CryRelatorio.Formulas(2) = "Fone='" & LcFone & "'"
 CryRelatorio.Formulas(3) = "titulo='Relatório de Comissões'"
 CryRelatorio.Formulas(7) = "Vendedor='" & Comissao.Text & "'"
 'CryRelatorio.Formulas(5) = "Versiculo1='" & LcVer1 & "'"
 CryRelatorio.Formulas(5) = "Celular='" & LcCelular & "'"
 CryRelatorio.Formulas(4) = "Celular1='" & Lccelular1 & "'"
 CryRelatorio.Formulas(6) = "email='" & Lcemail & "'"

If impressora Then
   LcTipoSaida = 1
Else
   LcTipoSaida = 0
End If
If analitico.Value = True Then CryRelatorio.SelectionFormula = LcFormula Else CryRelatorio.SelectionFormula = ""

CryRelatorio.Destination = LcTipoSaida
CryRelatorio.DiscardSavedData = True
CryRelatorio.PrintReport
Me.Caption = LcCap
'RsOpcao.Close
RsEmpresa.Close

Set RsOpcao = Nothing
Set RsEmpresa = Nothing

If CryRelatorio.LastErrorNumber > 0 Then
   MsgBox CryRelatorio.LastErrorString

End If
RsBaixa.Close
Set RsBaixa = Nothing
Dbbase.Close
End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub copias_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub

Private Sub Dataf_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Dataf_LostFocus()
If Not IsDate(Dataf.Text) And Dataf.Text <> "  /  /  " Then
      MsgBox "A data digitada não é Válida...", 48, "Aviso"
      Dataf.SetFocus
End If
End Sub

Private Sub Datai_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"
   
End Sub

Private Sub Datai_LostFocus()
If Not IsDate(Datai.Text) And Datai.Text <> "  /  /  " Then
      MsgBox "A data digitada não é Válida...", 48, "Aviso"
      Datai.SetFocus
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
DataS.Text = Format(GlDataSistema, "dd/mm/yyyy")
HoraS.Text = Format(Time, "hh:mm:ss")
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
LcIndice = "CODIGO"
Me.Height = 3675
Me.Width = 8685
CarregaVendedor

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
FrmPrincipal.SetFocus
End Sub

Private Sub Impressora_Click()
copias.Visible = True
Label3.Visible = True
End Sub
Function GeraSintetico()
Dim bb As Database
Dim LcComissao, LcTotal As Currency
Dim LcMuda, LcGrava As Integer
On Error GoTo EroGera
Set bb = OpenDatabase(GLBase, False, False) ' "dBASE III;")
LcCriterio1 = "Select * from alid201 where VENDEDOR='" & codigo.Text & "' and "
LcCriterio1 = LcCriterio1 & " DATAVENDA>=#" & Format(Datai.Text, "mm/dd/yyyy") & "# and DATAVENDA <=#" & Format(Dataf.Text, "mm/dd/yyyy") & "#"
LcCriterio1 = LcCriterio1 & " Order by Nf"
'MsgBox LcCriterio1

Set RsComissao = bb.OpenRecordset(LcCriterio1)
Set RsSintetico = bb.OpenRecordset("sintetico", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Do Until RsSintetico.EOF
   RsSintetico.Delete
   RsSintetico.MoveNext
Loop
LcNota = RsComissao!NF
Do Until RsComissao.EOF
   If LcMuda Then
      LcNota = RsComissao!NF
      LcMuda = False
   End If
   If LcNota = RsComissao!NF Then
      LcComissao = LcComissao + RsComissao!Comissao
      LcTotal = LcTotal + RsComissao!ValorTotal
      LcGrava = True
   Else
      RsComissao.MovePrevious
      Call GravaSintetico(LcComissao, LcTotal)
      LcComissao = 0
      LcTotal = 0
      LcMuda = True
      LcGrava = False
   End If
   RsComissao.MoveNext
   
Loop
If LcGrava Then
   RsComissao.MovePrevious
   Call GravaSintetico(LcComissao, LcTotal)
   LcGrava = False
End If
Exit Function
EroGera:
Resume Next
End Function
Function GravaSintetico(LcComissao, LcTotal As Currency)
Dim rsCliente As Recordset
Dim LcCliente As String
'On Error Resume Next
LcCriterio22 = "Select * from alid001 where codigo='" & RsComissao!Cliente & "'"
Set rsCliente = Dbbase.OpenRecordset(LcCriterio22)
If Not rsCliente.EOF Then
   LcCliente = rsCliente!RazaoSoc & ""
Else
   LcCliente = ""
End If
RsSintetico.AddNew
   RsSintetico!vendedor = RsComissao!vendedor
   RsSintetico!NF = RsComissao!NF
   RsSintetico!Comissao = LcComissao
   RsSintetico!ValorTotal = LcTotal
   
   RsSintetico!ItemBaixo = RsComissao!ItemBaixo
   RsSintetico!datavenda = RsComissao!datavenda
    RsSintetico!pago = RsComissao!pago
   RsSintetico!Cliente = LcCliente
RsSintetico.Update
rsCliente.Close
Set rsCliente = Nothing
End Function
Private Sub Impressora_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub

Private Sub marcar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub marcarpg_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub naoPago_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub pago_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub sintetico_Click()
'naoPago.Enabled = False
'pago.Enabled = False
'todos.Enabled = False

End Sub

Private Sub sintetico_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub

Private Sub todos_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Video_Click()
copias.Visible = False
Label3.Visible = False
End Sub

Private Sub Video_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub
