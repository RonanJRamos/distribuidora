VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form ProcessaBalanco 
   Caption         =   "Acerto Saldo"
   ClientHeight    =   1020
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10800
   Icon            =   "ProcessaBalanco.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1020
   ScaleWidth      =   10800
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Galpao 
      Height          =   315
      ItemData        =   "ProcessaBalanco.frx":0442
      Left            =   7920
      List            =   "ProcessaBalanco.frx":044F
      TabIndex        =   5
      Top             =   480
      Width           =   2175
   End
   Begin VB.CommandButton CmdImprimirRelatorio 
      Caption         =   "Imprimir Relatorios"
      Height          =   615
      Left            =   6480
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin Crystal.CrystalReport CryRelatorio 
      Left            =   120
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowAllowDrillDown=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.CommandButton CmdDadosGerados 
      Caption         =   "Gerar dados para o Relatorio"
      Height          =   615
      Left            =   5040
      TabIndex        =   3
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton CmdLancaZeros 
      Caption         =   "Lançar Produtos Zerados"
      Height          =   615
      Left            =   840
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton CmdSaldo 
      Caption         =   "Gerar Saldo"
      Height          =   615
      Left            =   3720
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton CmdPRocessar 
      Caption         =   "Processar"
      Height          =   615
      Left            =   2340
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Galpão a Imprimir"
      Height          =   255
      Left            =   7920
      TabIndex        =   6
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "ProcessaBalanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub ProcessaDadosRel(Datai As Date, dataf As Date, Mes As String)
Dim RsEstoque                   As ADODB.Recordset
Dim RsProduto                   As ADODB.Recordset
Dim RsEst                       As ADODB.Recordset
Dim Db                          As Database
Dim RsRel                       As Recordset
Dim SaldoSanta                  As Long
Dim SaldoSanta1                 As Long
Dim SaldoCalifornia             As Long

Dim SaldoSantaAnterior          As Long
Dim SaldoSanta1Anterior         As Long
Dim SaldoCaliforniaAnterior     As Long

Dim EntradasSanta               As Long
Dim EntradasSanta1              As Long
Dim EntradasCalifornia          As Long

Dim SaidasSanta                 As Long
Dim SaidasSanta1                As Long
Dim SaidasCalifornia            As Long

Dim Atual                       As Long
Dim TotalReg                    As Long
Dim StrSql                      As String
Dim Posicao                     As Long
Dim Unidade                     As String

LcCap = Me.Caption
Screen.MousePointer = 11
'==> Abre tabela de estoque fiscal
Set Db = OpenDatabase(GLBase)
'Db.Execute "Delete from relestoquefiscal"
Set RsRel = Db.OpenRecordset("Select * from relestoquefiscal")
StrSql = "Select * from estoquefiscal where data between '" & Format(Datai, "yyyy-mm-dd") & "' and '" & Format(dataf, "yyyy-mm-dd") & "' order by posicao;"

'MsgBox StrSql
Set RsEstoque = AbreRecordset(StrSql, RsEstoque)
If Not RsEstoque.EOF Then
   RsEstoque.MoveLast
  ' MsgBox RsEstoque.RecordCount
   RsEstoque.MoveFirst
End If
'==> Abre tabela Produtos
StrSql = "Select * from Produtos order by codigo"
Set RsProduto = AbreRecordsetLeitura(StrSql)

If Not RsProduto.EOF Then
   RsProduto.MoveLast
   TotalReg = RsProduto.RecordCount
   RsProduto.MoveFirst
End If
Atual = 0


Do Until RsProduto.EOF
    Atual = Atual + 1
    Me.Caption = "Processando Reg " & Atual & " de " & TotalReg & " " & RsProduto!Nome
    DoEvents
    SaldoAnterior = 0
    SaldoSanta = 0
    SaldoSanta1 = 0
    SaldoCalifornia = 0
    SaldoGeral = 0
    EntradasSanta = 0
    EntradasSanta1 = 0
    EntradasCalifornia = 0
    SaidasSanta = 0
    SaidasSanta1 = 0
    SaidasCalifornia = 0
    SaldoSantaAnterior = 0
    SaldoSanta1Anterior = 0
    SaldoCaliforniaAnterior = 0
    Posicao = 0
    RsEstoque.Filter = "CodigoPRoduto=" & RsProduto!Codigo
   ' RsEstoque.Sort = "Posicao"
    Posicao = 1
    '==> Posiciona o proximo registro.
    If RsEstoque.EOF Then
            StrSql = "Select * from estoquefiscal where data > '" & Format(dataf, "yyyy-mm-dd") & "' and codigoproduto=" & RsProduto!Codigo & " order by posicao desc;"
            'Debug.Print StrSql
            'MsgBox StrSql
            Set RsEst = AbreRecordsetLeitura(StrSql)
            If Not RsEst.EOF Then
                SaldoSantaAnterior = RsEst!SaldoAnteriorSanta
                SaldoSanta1Anterior = RsEst!SaldoAnteriorSanta1
                SaldoCaliforniaAnterior = RsEst!SaldoAnteriorCalifornia
                
                SaldoSanta = RsEst!SaldoAnteriorSanta
                SaldoSanta1 = RsEst!SaldoAnteriorSanta1
                SaldoCalifornia = RsEst!SaldoAnteriorCalifornia
                EntradasSanta = 0
                EntradasSanta1 = 0
                EntradasCalifornia = 0
                SaidasSanta = 0
                SaidasSanta1 = 0
                SaidasCalifornia = 0
            End If
    End If
    Do Until RsEstoque.EOF
       Unidade = RsEstoque!Unidade
       If Posicao = 1 Then
            SaldoSantaAnterior = RsEstoque!SaldoAnteriorSanta
            SaldoSanta1Anterior = RsEstoque!SaldoAnteriorSanta1
            SaldoCaliforniaAnterior = RsEstoque!SaldoAnteriorCalifornia
            
            SaldoSanta = RsEstoque!SaldoSanta
            SaldoSanta1 = RsEstoque!SaldoSanta1
            SaldoCalifornia = RsEstoque!SaldoCalifornia
            EntradasSanta = EntradasSanta + RsEstoque!QTEntradaSanta
            EntradasSanta1 = EntradasSanta1 + RsEstoque!QTEntradaSanta1
            EntradasCalifornia = EntradasCalifornia + RsEstoque!QTEntradaCalifornia
            SaidasSanta = SaidasSanta + RsEstoque!QTSaidaSanta
            SaidasSanta1 = SaidasSanta1 + RsEstoque!QTSaidaSanta1
            SaidasCalifornia = SaidasCalifornia + RsEstoque!QTSaidaCalifornia
       Else
            SaldoSanta = RsEstoque!SaldoSanta
            SaldoSanta1 = RsEstoque!SaldoSanta1
            SaldoCalifornia = RsEstoque!SaldoCalifornia
            EntradasSanta = EntradasSanta + RsEstoque!QTEntradaSanta
            EntradasSanta1 = EntradasSanta1 + RsEstoque!QTEntradaSanta1
            EntradasCalifornia = EntradasCalifornia + RsEstoque!QTEntradaCalifornia
            SaidasSanta = SaidasSanta + RsEstoque!QTSaidaSanta
            SaidasSanta1 = SaidasSanta1 + RsEstoque!QTSaidaSanta1
            SaidasCalifornia = SaidasCalifornia + RsEstoque!QTSaidaCalifornia
       End If
       RsEstoque.MoveNext
    Loop
    
    RsRel.AddNew
    RsRel!CodigoPRoduto = RsProduto!Codigo
    RsRel!Nome = RsProduto!Nome
    RsRel!Unidade = Unidade
    RsRel!QASanta = SaldoSantaAnterior
    RsRel!QASanta1 = SaldoSanta1Anterior
    RsRel!QACalifornia = SaldoCaliforniaAnterior
    RsRel!QSSanta = SaidasSanta
    RsRel!QSSanta1 = SaidasSanta1
    RsRel!QSCalifornia = SaidasCalifornia
    
    RsRel!QESanta = EntradasSanta
    RsRel!QESanta1 = EntradasSanta1
    RsRel!QECalifornia = EntradasCalifornia
    RsRel!SaldoSanta = SaldoSanta
    RsRel!SaldoSanta1 = SaldoSanta1
    RsRel!SaldoCalifornia = SaldoCalifornia
    RsRel!Mes = Mes
    RsRel.Update
    
    RsProduto.MoveNext
Loop
Me.Caption = LcCap
Screen.MousePointer = 0

End Sub

Private Sub CmdDadosGerados_Click()
 Dim Db                          As Database
 
 Set Db = OpenDatabase(GLBase)
 'Db.Execute ("Delete from estoquefiscal")
 Db.Execute "Delete from relestoquefiscal"
 afetados = Db.RecordsAffected
'==> Janeiro
ProcessaDadosRel CDate("01/01/06"), CDate("31/01/06"), "Janeiro"
'==> Fevereiro
ProcessaDadosRel CDate("01/02/06"), CDate("28/02/06"), "Fevereiro"
'==> março
ProcessaDadosRel CDate("01/03/06"), CDate("31/03/06"), "Março"
'==> Abril
ProcessaDadosRel CDate("01/04/06"), CDate("30/04/06"), "Abril"
'==> Maio
ProcessaDadosRel CDate("01/05/06"), CDate("31/05/06"), "Maio"
'==> Junho
ProcessaDadosRel CDate("01/06/06"), CDate("30/06/06"), "Junho"
'==> Julho
ProcessaDadosRel CDate("01/07/06"), CDate("31/07/06"), "Julho"
'==> Agosto
ProcessaDadosRel CDate("01/08/06"), CDate("31/08/06"), "Agosto"
MsgBox "terminado"

End Sub

Private Sub CmdImprimirRelatorio_Click()
On Error Resume Next

'Abertura do relatório de vendas
 ImprimeRel "Janeiro"
 ImprimeRel "Fevereiro"
 ImprimeRel "Março"
 ImprimeRel "Abril"
 ImprimeRel "Maio"
 ImprimeRel "Junho"
 ImprimeRel "Julho"
 ImprimeRel "Agosto"

End Sub
Sub ImprimeRel(Mes As String)
On Error Resume Next

'Abertura do relatório de vendas
    
    
CryRelatorio.DataFiles(0) = GLBase

Select Case UCase(Galpao.Text)

    Case Is = UCase("Santa Maria")
        CryRelatorio.ReportFileName = App.Path & "\InventarioSanta.rpt"
    Case Is = UCase("Santa Maria 1")
        CryRelatorio.ReportFileName = App.Path & "\InventarioSanta1.rpt"
    Case Is = UCase("California")
        CryRelatorio.ReportFileName = App.Path & "\InventarioCalifornia.rpt"
    Case Else
        MsgBox "Escolha um galpão valido para a impressão."
        Galpao.SetFocus
        Exit Sub
End Select
CryRelatorio.CopiesToPrinter = Val(copias.Text)
    
CryRelatorio.DiscardSavedData = True
CryRelatorio.WindowTop = 50
CryRelatorio.WindowWidth = 700
CryRelatorio.WindowLeft = 50
CryRelatorio.WindowHeight = 500
CryRelatorio.WindowTitle = "Invetário de Estoque."

LcFormula = "{RelEstoqueFiscal.Mes}='" & Mes & "'"
'CryRelatorio.SelectionFormula = LcFormula
CryRelatorio.SelectionFormula = LcFormula
CryRelatorio.Destination = 0
CryRelatorio.PrintReport
'RsOpcao.Close

If CryRelatorio.LastErrorNumber > 0 Then MsgBox CryRelatorio.LastErrorString

End Sub
Private Sub CmdLancaZeros_Click()
Dim RsEstoque                   As ADODB.Recordset
Dim RsProduto                   As ADODB.Recordset
Dim Unidade                     As String
Dim StrSql                      As String
Dim Db                          As Database
Dim RsUnbidade                  As Recordset

LcCap = Me.Caption
Screen.MousePointer = 11
'==> Abre tabela de estoque fiscal
StrSql = "Select * from estoquefiscal where data between '2006-01-01' and '2006-09-01' order by data desc,codigo asc;"

'MsgBox StrSql
'Debug.Print StrSql
Set RsEstoque = AbreRecordset(StrSql, RsEstoque)

'==> Abre tabela Produtos
StrSql = "Select * from Produtos order by codigo"
Set RsProduto = AbreRecordsetLeitura(StrSql)

If Not RsProduto.EOF Then
   RsProduto.MoveLast
   TotalReg = RsProduto.RecordCount
   RsProduto.MoveFirst
End If
Atual = 0
Set Db = OpenDatabase(GLBase)
Set RsUnidade = Db.OpenRecordset("Select * from alid004 order by cod")

Do Until RsProduto.EOF
    Atual = Atual + 1
    Me.Caption = "Processando Reg " & Atual & " de " & TotalReg & " " & RsProduto!Nome
    DoEvents
    
    RsEstoque.Filter = "CodigoPRoduto=" & RsProduto!Codigo
    If RsEstoque.EOF Then
    
        RsUnidade.FindFirst "Cod='" & RsProduto!unidmedida & "'"
        Unidade = ""
        If Not RsUnidade.NoMatch Then
            Unidade = RsUnidade!Simbolo & ""
        End If
        StrSql = "Insert into estoquefiscal(data,codigoproduto,unidade,quantidadeEntradageral," & _
             "valorcustomediounitario,vcustototal,saldoGeral,QuantidadeSaidaGeral,QTEntradaSanta," & _
             "QTEntradaSanta1,QTEntradaCalifornia,QTSaidaSanta,QTSaidaSanta1,QTSaidaCalifornia," & _
             "Saldosanta,Saldosanta1,SaldoCalifornia,SaldoAnteriorSanta,SaldoAnteriorSanta1,SaldoAnteriorCalifornia,nome) Values ('" & _
             Format("01/09/06", "yyyy-mm-dd") & "'," & _
             RsProduto!Codigo & ",'" & _
             Unidade & "'," & _
             "0" & "," & _
             "0" & "," & _
             "0" & "," & _
             "0" & "," & _
             "0" & "," & _
             "0" & "," & _
             "0" & "," & _
             "0" & "," & _
             "0" & "," & _
             "0" & "," & _
             "0" & "," & _
             "0" & "," & _
             "0" & "," & _
             "0" & "," & _
             "0" & "," & _
             "0" & "," & _
             "0" & ",'" & _
             Replace(Replace(Replace(RsProduto!Nome, ",", ""), "'", ""), Chr(34), "") & "')"
             'Debug.Print StrSql
             'MsgBox StrSql
        conexaoAdo.Execute StrSql, afetados
    End If
    
    RsProduto.MoveNext
Loop

Me.Caption = LcCap
Screen.MousePointer = 0

Set RsProduto = Nothing
Set RsEstoque = Nothing

MsgBox "Processo terminado"
End Sub

Private Sub CmdPRocessar_Click()
Dim StrSql As String
Dim Db            As Database
Dim RsUnidade     As Recordset
Dim RsEntrada     As ADODB.Recordset
Dim RsSaida       As ADODB.Recordset
Dim RsItemSaida   As ADODB.Recordset
Dim RsProduto     As ADODB.Recordset
Dim SaldoAnterior As Double
Dim Saldo         As Double
Dim TotalRegistro As Long
Dim Atual         As Long
Dim LcCap         As String
Dim Unidade       As String
Dim Quantidade    As Long
Dim QtSanta1      As Long
Dim QtCali        As Long

Set Db = OpenDatabase(GLBase)
CmdPRocessar.Enabled = False
DoEvents
Set RsUnidade = Db.OpenRecordset("Select * from alid004")
LcCap = Me.Caption
Screen.MousePointer = 11
'==> Lanca as Notas de entrada

StrSql = "select * from itensentradanf where data between '2006-01-01' and '2006-08-31' order by data;"
Set RsEntrada = AbreRecordsetLeitura(StrSql)
StrSql = "Select * from alid052 order by numnf;"
Set RsItemSaida = AbreRecordsetLeitura(StrSql)
StrSql = "Select * from produtos order by codigo"
Set RsProduto = AbreRecordsetLeitura(StrSql)

'If Not RsEntrada.EOF Then
'   RsEntrada.MoveLast
'   TotalRegistro = RsEntrada.RecordCount
'   RsEntrada.MoveFirst
'End If

'Atual = 0
'Do Until RsEntrada.EOF
'  Atual = Atual + 1
'  Me.Caption = "Processando entrada " & RsEntrada!numnf & " data " & Format(RsEntrada!Data, "dd/mm/yy") & " Reg:" & Atual & " de " & TotalRegistro
'  DoEvents
'  RsUnidade.FindFirst "Cod='" & RsEntrada!Unimed & "'"
'  Unidade = ""
'  If Not RsUnidade.NoMatch Then
'     Unidade = RsUnidade!Simbolo & ""
'  End If
'  StrSql = "Insert into estoquefiscal(data,codigoproduto,unidade,quantidadeEntradageral," & _
         "valorcustomediounitario,vcustototal,saldoGeral,QuantidadeSaidaGeral,QTEntradaSanta," & _
         "QTEntradaSanta1,QTEntradaCalifornia,QTSaidaSanta,QTSaidaSanta1,QTSaidaCalifornia," & _
         "Saldosanta,Saldosanta1,SaldoCalifornia,nome) Values ('" & _
         Format(RsEntrada!Data, "yyyy-mm-dd") & "'," & _
         RsEntrada!Item & ",'" & _
         Unidade & "'," & _
         "0" & "," & _
         Replace(RsEntrada!VALUNIT, ",", ".") & "," & _
         Replace(RsEntrada!ValorTotal, ",", ".") & "," & _
         "0" & "," & _
         "0" & "," & _
         Replace(RsEntrada!qtde, ",", ".") & "," & _
         "0" & "," & _
         "0" & "," & _
         "0" & "," & _
         "0" & "," & _
         "0" & "," & _
         "0" & "," & _
         "0" & "," & _
         "0" & ",'" & _
         Replace(Replace(Replace(RsEntrada!Descricao, ",", ""), "'", ""), Chr(34), "") & "')"
   ' Debug.Print StrSql
   ' MsgBox StrSql
'    conexaoAdo.Execute StrSql, Afetados
'  RsEntrada.MoveNext
'Loop

'Set RsEntrada = Nothing

'==> Lanca as Notas de Saida

StrSql = "select * from alid050 where dtemis between '2006-01-01' and '2006-08-31' order by numnf;"

Set RsSaida = AbreRecordsetLeitura(StrSql)

If Not RsSaida.EOF Then
   RsSaida.MoveLast
   TotalRegistro = RsSaida.RecordCount
   RsSaida.MoveFirst
End If

Atual = 0
Do Until RsSaida.EOF
  Atual = Atual + 1
  Me.Caption = "Processando Saida " & RsSaida!numnf & " data " & Format(RsSaida!dtemis, "dd/mm/yy") & " Reg:" & Atual & " de " & TotalRegistro
  DoEvents
   '==> Separando os itens da nota
  RsItemSaida.Filter = ""
  RsItemSaida.Filter = "Numnf='" & RsSaida!numnf & "'"
  If Not RsItemSaida.EOF Then
    RsProduto.Filter = "Codigo=" & RsItemSaida!codprod
    If Not RsProduto.EOF Then
        '==> Verifica se as unidades são iguais
        RsUnidade.FindFirst "cod='" & RsProduto!unidmedida & "'"
        If Not RsUnidade.NoMatch Then
           '==> verifica se as unidades sao iguais
           If RsUnidade!Simbolo = RsItemSaida!Unimed Then
              Quantidade = RsItemSaida!qtde
           Else
             '==> Sao diferentes
             Quantidade = RsItemSaida!qtde * RsItemSaida!qtdum
             Quantidade = Quantidade / IIf(RsProduto!QtdMedida > 0, RsProduto!QtdMedida, 1)
           End If
        Else
           Quantidade = RsItemSaida!qtde
        End If
        QtSanta1 = 0
        QtCali = 0
        If RsSaida!Cliente = "03916" Then QtSanta1 = Quantidade
        If RsSaida!Cliente = "03915" Then QtCali = Quantidade
        
        Unidade = RsItemSaida!Unimed & ""
        StrSql = "Insert into estoquefiscal(data,codigoproduto,unidade,quantidadeEntradageral," & _
               "valorcustomediounitario,vcustototal,saldoGeral,QuantidadeSaidaGeral,QTEntradaSanta," & _
               "QTEntradaSanta1,QTEntradaCalifornia,QTSaidaSanta,QTSaidaSanta1,QTSaidaCalifornia," & _
               "Saldosanta,Saldosanta1,SaldoCalifornia,nome) Values ('" & _
               Format(RsSaida!dtemis, "yyyy-mm-dd") & "'," & _
               RsItemSaida!codprod & ",'" & _
               Unidade & "'," & _
               "0" & "," & _
               "0" & "," & _
               "0" & "," & _
               "0" & "," & _
               "0" & "," & _
               "0" & "," & _
               Replace(QtSanta1, ",", ".") & "," & _
               Replace(QtCali, ",", ".") & "," & _
               Replace(Quantidade, ",", ".") & "," & _
               "0" & "," & _
               "0" & "," & _
               "0" & "," & _
               "0" & "," & _
               "0" & ",'" & _
               Replace(Replace(Replace(RsItemSaida!Descricao, ",", ""), "'", ""), Chr(34), "") & "')"
               
          conexaoAdo.Execute StrSql, afetados
      End If
  End If
  RsSaida.MoveNext
Loop

Me.Caption = LcCap
Screen.MousePointer = 0

MsgBox "Processo terminado."
CmdPRocessar.Enabled = True
End Sub

Private Sub CmdSaldo_Click()
Dim RsEstoque                   As ADODB.Recordset
Dim RsProduto                   As ADODB.Recordset
Dim SaldoAnterior               As Long
Dim SaldoSanta                  As Long
Dim SaldoSanta1                 As Long
Dim SaldoCalifornia             As Long

Dim SaldoSantaAnterior          As Long
Dim SaldoSanta1Anterior         As Long
Dim SaldoCaliforniaAnterior     As Long

Dim SaldoGeral                  As Long
Dim Atual                       As Long
Dim TotalReg                    As Long
Dim StrSql                      As String
Dim Posicao                     As Long

LcCap = Me.Caption
Screen.MousePointer = 11
'==> Abre tabela de estoque fiscal
StrSql = "Select * from estoquefiscal where data between '2006-01-01' and '2006-09-01' order by data desc,codigo asc;"

Set RsEstoque = AbreRecordsetLeitura(StrSql)

'==> Abre tabela Produtos
StrSql = "Select * from Produtos order by codigo"
Set RsProduto = AbreRecordsetLeitura(StrSql)

If Not RsProduto.EOF Then
   RsProduto.MoveLast
   TotalReg = RsProduto.RecordCount
   RsProduto.MoveFirst
End If
Atual = 0


Do Until RsProduto.EOF
    Atual = Atual + 1
    Me.Caption = "Processando Reg " & Atual & " de " & TotalReg & " " & RsProduto!Nome
    DoEvents
    SaldoAnterior = 0
    SaldoSanta = 0
    SaldoSanta1 = 0
    SaldoCalifornia = 0
    SaldoGeral = 0
    
    SaldoSantaAnterior = 0
    SaldoSanta1Anterior = 0
    SaldoCaliforniaAnterior = 0
    SaldoGeralAnterior = 0
    'If RsProduto!Codigo = 16 Then Stop
    RsEstoque.Filter = "CodigoPRoduto=" & RsProduto!Codigo
   ' StrSql = "Select * from estoquefiscal where data between '2006-01-01' and '2006-09-01' and CodigoPRoduto=" & RsProduto!Codigo & " order by data desc,codigo asc;"

   ' Set RsEstoque = AbreRecordset(StrSql, RsEstoque)

    Posicao = 1
    Do Until RsEstoque.EOF
       If CDate(RsEstoque!Data) = CDate("01/09/06") Then
          SaldoAnterior = 0
          SaldoSanta = RsEstoque!SaldoSanta
          SaldoSanta1 = RsEstoque!SaldoSanta1
          SaldoCalifornia = RsEstoque!SaldoCalifornia
          SaldoGeral = SaldoSanta + SaldoSanta1 + SaldoCalifornia
            
          SaldoSantaAnterior = SaldoSanta
          SaldoSanta1Anterior = SaldoSanta1
          SaldoCaliforniaAnterior = SaldoCalifornia
          SaldoGeralAnterior = SaldoGeral
          
          'RsEstoque!QTEntradaSanta = 0
          'RsEstoque!QTEntradaSanta1 = 0
          'RsEstoque!QTEntradaCalifornia = 0
          
          'RsEstoque!SaldoAnteriorSanta = SaldoSantaAnterior
          'RsEstoque!SaldoAnteriorSanta1 = SaldoSanta1Anterior
          'RsEstoque!SaldoAnteriorCalifornia = SaldoCaliforniaAnterior
          'RsEstoque!Posicao = Posicao
          'RsEstoque.Update
          StrSql = "Update estoquefiscal set " & _
                   "QTEntradaSanta=0 " & _
                   ",QTEntradaSanta1 = 0" & _
                   ",QTEntradaCalifornia = 0" & _
                   ",SaldoAnteriorSanta = " & Replace(SaldoSantaAnterior, ",", ".") & _
                   ",SaldoAnteriorSanta1 = " & Replace(SaldoSanta1Anterior, ",", ".") & _
                   ",SaldoAnteriorCalifornia = " & Replace(SaldoCaliforniaAnterior, ",", ".") & _
                   ",Posicao =" & Posicao & _
                   " Where codigo=" & RsEstoque!Codigo
          conexaoAdo.Execute StrSql, afetados
          
       Else
          SaldoSanta = SaldoSantaAnterior
          SaldoSanta1 = SaldoSanta1Anterior
          SaldoCalifornia = SaldoCaliforniaAnterior
          SaldoGeral = SaldoGeralAnterior
        
          SaldoSantaAnterior = SaldoSanta + RsEstoque!QTSaidaSanta - RsEstoque!QTEntradaSanta
          SaldoSanta1Anterior = SaldoSanta1 + RsEstoque!QTSaidaSanta1 - RsEstoque!QTEntradaSanta1
          SaldoCaliforniaAnterior = SaldoCalifornia + RsEstoque!QTEntradaCalifornia - RsEstoque!QTEntradaCalifornia
          SaldoGeralAnterior = SaldoSanta + SaldoSanta1 + SaldoCalifornia
          
          'RsEstoque!SaldoSanta = SaldoSanta
          'RsEstoque!SaldoSanta1 = SaldoSanta1
          'RsEstoque!SaldoCalifornia = SaldoCalifornia
          'RsEstoque!SaldoGeral = SaldoGeral
          
          'RsEstoque!SaldoAnteriorSanta = SaldoSantaAnterior
          'RsEstoque!SaldoAnteriorSanta1 = SaldoSanta1Anterior
          'RsEstoque!SaldoAnteriorCalifornia = SaldoCaliforniaAnterior
          'RsEstoque!Posicao = Posicao
          'RsEstoque.Update
          
           StrSql = "Update estoquefiscal set " & _
                   "SaldoSanta=" & Replace(SaldoSanta, ",", ".") & _
                   ",SaldoSanta1 =" & Replace(SaldoSanta1, ",", ".") & _
                   ",SaldoCalifornia =" & Replace(SaldoCalifornia, ",", ".") & _
                   ",SaldoGeral =" & Replace(SaldoGeral, ",", ".") & _
                   ",SaldoAnteriorSanta = " & Replace(SaldoSantaAnterior, ",", ".") & _
                   ",SaldoAnteriorSanta1 = " & Replace(SaldoSanta1Anterior, ",", ".") & _
                   ",SaldoAnteriorCalifornia = " & Replace(SaldoCaliforniaAnterior, ",", ".") & _
                   ",Posicao =" & Posicao & _
                   " Where codigo=" & RsEstoque!Codigo
          conexaoAdo.Execute StrSql, afetados
          
          
       End If
       Posicao = Posicao + 1
       RsEstoque.MoveNext
    Loop

  RsProduto.MoveNext
Loop

Me.Caption = LcCap
Screen.MousePointer = 0

Set RsProduto = Nothing
Set RsEstoque = Nothing

MsgBox "Processo terminado"



End Sub
