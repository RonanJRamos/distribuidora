VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmInventarioCusto 
   BackColor       =   &H00C00000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inventario com Custo medio"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4455
   ForeColor       =   &H00000000&
   Icon            =   "FrmInventarioCusto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C00000&
      Caption         =   "Considerar pela Unidade cadastrada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   2280
      Width           =   3975
   End
   Begin VB.TextBox Produto 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exibir Relatorio"
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   4095
   End
   Begin MSMask.MaskEdBox DataI 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox DataF 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "FrmInventarioCusto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Rel As New CrysInventarioComCusto

Private Sub Command1_Click()

'On Error Resume Next
Dim StrSql          As String
Dim Rs As ADODB.Recordset
LcCap = Me.Caption
Me.Caption = "Aguarde, processando os dados..."
Screen.MousePointer = 11
GeraDados
Screen.MousePointer = 0
Me.Caption = LcCap
Screen.MousePointer = vbHourglass

StrSql = "Select * from relatorioinvfiscal order by nome"

Set Rs = AbreRecordset(StrSql, True)
Load Relatorios
With Relatorios
     Rel.DiscardSavedData
     Rel.Database.SetDataSource Rs
     .CRViewer1.ReportSource = Rel
    ' setaformula
      .CRViewer1.ViewReport
End With
Relatorios.Show
Screen.MousePointer = vbDefault
End Sub

Sub GeraDados()
On Error GoTo errGera
Dim RsEntrada As ADODB.Recordset
Dim RsSaida As ADODB.Recordset
Dim RsRel   As ADODB.Recordset
Dim StrSql As String
Dim Quantidade As Double
Dim SaldoAnterior As Double
Dim LcSaldoTotal As Single

'==> Busca a Entrada
StrSql = "SELECT Sum(itensentradanf.QTDE) AS SomaDeQTDE, itensentradanf.ITEM, produtos.NOME, produtos.QtdMedida " & _
         "FROM (entradanf INNER JOIN itensentradanf ON entradanf.codigo = itensentradanf.CodigoNota) INNER JOIN produtos ON itensentradanf.ITEM = produtos.codigo " & _
         "WHERE (((entradanf.DATA) Between #" & Format(Datai.Text, "mm/dd/yy") & "# And #" & Format(Dataf.Text, "mm/dd/yy") & "#) AND ((itensentradanf.descricao) Like '" & Produto.Text & "%')) " & _
         "GROUP BY itensentradanf.ITEM, produtos.NOME;"

Set RsEntrada = AbreRecordset(StrSql, True)
Debug.Print StrSql

'==> Busca as saidas
StrSql = "SELECT produtos.codigo, produtos.NOME, alid052.QTDE, alid052.QTDUM, produtos.QtdMedida " & _
         "FROM (alid050 INNER JOIN alid052 ON alid050.NUMNF = alid052.NUMNF) INNER JOIN produtos ON alid052.codProd = produtos.codigo " & _
         "WHERE (((alid050.DTEMIS) Between #" & Format(Datai.Text, "mm/dd/yy") & "# And #" & Format(Dataf.Text, "mm/dd/yy") & "#) AND ((alid052.descricao) Like '" & Produto.Text & "%')) " & _
         "GROUP BY produtos.codigo, produtos.NOME, alid052.QTDUM, produtos.QtdMedida " & _
         "ORDER BY produtos.codigo;"
         
Set RsSaida = AbreRecordset(StrSql, True)
'MsgBox DEscricaoErro
'==> Esclui a Tabela
afetados = ExecutaSql("Delete from relatorioinvfiscal")
'==> Inclui os dados da entrada
'Debug.Print StrSql
Do Until RsEntrada.EOF
  Dim LcCusto As Single
  LcCusto = 0
  LcCusto = CalculaCustomedio(RsEntrada!Item)
  SaldoAnterior = 0
  LcQuantidade = 0
  If Check1.Value = 0 Then
     LcQuantidade = RsEntrada!SomaDeQTDE * RsEntrada!QtdMedida
  Else
    LcQuantidade = RsEntrada!SomaDeQTDE
  End If
  
  'SaldoAnterior = BuscaSaldoAnteriorProduto(RsEntrada!Item)
  StrSql = "Insert into relatorioinvfiscal (CodProd,Nome,Entrada,Saidas,Anterior,CustoMedio) vALUES ("
  StrSql = StrSql & RsEntrada!Item & ",'"
  StrSql = StrSql & Replace(RsEntrada!Nome, "'", "''") & "',"
  StrSql = StrSql & Replace(LcQuantidade, ",", ".") & ","
  StrSql = StrSql & "0,"
  StrSql = StrSql & Replace(SaldoAnterior, ",", ".") & ","
  StrSql = StrSql & Replace(LcCusto, ",", ".") & ")"
  afetados = ExecutaSql(StrSql)
 ' MsgBox DEscricaoErro
 ' MsgBox StrSql
  RsEntrada.MoveNext
Loop
Do Until RsSaida.EOF
  '==> Verifica se ja existe o produto
  StrSql = "Select * from relatorioinvfiscal where CodProd=" & RsSaida!codigo
  Set RsRel = AbreRecordset(StrSql, True)
  'MsgBox DEscricaoErro
  Quantidade = 0
  If Check1.Value = 0 Then
     Quantidade = (RsSaida!Qtde * RsSaida!QTDUM) ' / IIf(RsSaida!QtdMedida > 0, RsSaida!QtdMedida, 1)
  Else
     Quantidade = (RsSaida!Qtde * RsSaida!QTDUM) / IIf(RsSaida!QtdMedida > 0, RsSaida!QtdMedida, 1)
  End If
  If Not RsRel.EOF Then
     StrSql = "Update relatorioinvfiscal set Saidas=" & Replace(CStr(RsRel!Saidas + Quantidade), ",", ".")
     StrSql = StrSql & " Where codigo=" & RsRel!codigo
  Else
     SaldoAnterior = 0
   ' SaldoAnterior = BuscaSaldoAnteriorProduto(RsSaida!codigo)

    StrSql = "Insert into relatorioinvfiscal (CodProd,Nome,Entrada,Saidas,Anterior) vALUES ("
    StrSql = StrSql & RsSaida!codigo & ",'"
    StrSql = StrSql & Replace(RsSaida!Nome, "'", "''") & "',"
    StrSql = StrSql & "0,"
    StrSql = StrSql & Replace(CStr(Quantidade), ",", ".") & ","
    StrSql = StrSql & Replace(SaldoAnterior, ",", ".") & ")"
  End If
  afetados = ExecutaSql(StrSql)
  'MsgBox StrSql
 ' MsgBox DEscricaoErro
  RsSaida.MoveNext
Loop
'==> Calcula o custo total
StrSql = "Select * from relatorioinvfiscal"
Set RsRel = AbreRecordset(StrSql)
Do Until RsRel.EOF
   Lcsaldo = RsRel!Entrada - RsRel!Saidas
   LcSaldoTotal = RsRel!CustoMedio * Lcsaldo
   StrSql = "Update relatorioinvfiscal set CustomedioTotal=" & Replace(CStr(LcSaldoTotal), ",", ".") & ",saldo=" & Replace(CStr(Lcsaldo), ",", ".") & " where CodProd=" & RsRel!codProd
   afetados = ExecutaSql(StrSql)
   
   'MsgBox StrSql
   RsRel.MoveNext
Loop
Exit Sub
errGera:
MsgBox err.Description & "  " & err.Number
Resume 0
End Sub
Function CalculaCustomedio(codigoproduto) As Single
Dim Resposta As Single
Dim LcCusto As Single
Dim LcCom As Single
Dim StrSql As String
Dim RsCusto As ADODB.Recordset
Dim a As Long
StrSql = "Select * from itensentradanf where ITEM='" & codigoproduto & "'"
Set RsCusto = AbreRecordset(StrSql, True)
a = 0
Resposta = 0
Do Until RsCusto.EOF
  a = a + 1
  If IsNumeric(RsCusto!QTDUM) Then LcCom = RsCusto!QTDUM Else LcCom = 1
  If LcCom = 0 Then LcCom = 1
  If Check1.Value = 0 Then
     LcCusto = RsCusto!VALUNIT / LcCom
  Else
    LcCusto = RsCusto!VALUNIT
  End If
  Resposta = (Resposta + LcCusto) / a
  RsCusto.MoveNext
Loop
CalculaCustomedio = Resposta
Set RsCusto = Nothing

End Function
