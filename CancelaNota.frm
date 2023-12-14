VERSION 5.00
Begin VB.Form CancelaNota 
   BackColor       =   &H00E0F8FC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cancelamento da Nota de Saída"
   ClientHeight    =   1185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar F10"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Confirma F3"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox NF 
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
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
   Begin VB.Line Line1 
      X1              =   2640
      X2              =   2640
      Y1              =   0
      Y2              =   1200
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Número da Nota"
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
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "CancelaNota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo ErrCancelamento
Dim db As Database
Dim LcSql As String
'Dim RsNota As ADODB.Recordset
Dim RsItens As ADODB.Recordset
Dim RsHistorico As ADODB.Recordset
Dim RsNota As ADODB.Recordset
Dim rsCliente As Recordset

Dim Cestoque As New ControleDb
Dim LcSanta As Double
Dim LcSanta1 As Double
Dim LcCalifornia As Double
If Len(NF.Text) = 0 Then
   MsgBox "Digite o Número da Nota fiscal.", 64, "Aviso"
   NF.SetFocus
   Exit Sub
End If
If Left(UCase(NF.Text), 2) = "AL" Then
   CancelaAL
   Exit Sub
End If
NF.Text = Right("000000" & NF.Text, 6)
conexaoAdo.BeginTrans
Set db = OpenDatabase(GLBase)
LcCap = Me.Caption
Me.Caption = "Aguarde, Desfazendo Lançamentos..."
Set RsNota = AbreRecordset("Select * from alid050 where numnf='" & NF.Text & "'", True)

If RsNota.EOF Then
   MsgBox "A Nota Fiscal não foi encontrada.", 64, "Aviso"
   conexaoAdo.CommitTrans
   GoTo Saida
End If
If RsNota("Status") = "CANCELADA" Then
   MsgBox "Esta nota já foi cancelada.", 64, "Aviso"
   conexaoAdo.CommitTrans
   GoTo Saida
End If
Set RsItens = AbreRecordset("Select * from alid052 where numnf='" & NF.Text & "'", True)
Me.Caption = "Marcando Nota Fiscal como Cancelada."
Me.Caption = "Excluido comissões..."
LcSql = "Delete from alid201 where nf='" & NF.Text & "'"
db.Execute LcSql
Me.Caption = "Acertando dados do Sintegra..."
LcSql = "UPDATE sintegra_50 Set situacao='S' where nf='" & NF.Text & "' and cfop='" & Replace(RsNota!CFOP, ".", "") & "'"
afetados = ExecutaSql(LcSql)

Me.Caption = "Excluido Contas a Receber do Cliente..."
LcSql = "Delete from alid015 where nf like '" & NF.Text & "%'"
ExecutaSql LcSql
'==> Estorna Credito Cliente

Set RsNota = AbreRecordset("Select * from Alid050 where numnf='" & NF.Text & "'", True)
If Not RsNota.EOF Then
   Set rsCliente = db.OpenRecordset("Select * from Alid001 where codigo='" & Right("00000" & RsNota!Cliente, 5) & "'")
    If UCase(RsNota!Natureza) <> UCase("VV") And UCase(RsNota!Natureza) <> "DE" And UCase(RsNota!Natureza) <> "TR" And UCase(RsNota!Natureza) <> "RE" Then
        'LcCriterioPes = "codigo='" & txt(8).Text & "'"
        'RsCliente.FindFirst LcCriterioPes
        If Not rsCliente.EOF Then
           rsCliente.Edit
         '  rsCliente("ULTCOMPRA") = CDate(txt(12).Text)
           'If Natureza.Text <> "DEVOLUCAO" And Natureza.Text <> "TRANSFERENCIA" Then
              rsCliente("CreditoUtilizado") = rsCliente("CreditoUtilizado") - CCur(RsNota!ValorNota)
           'End If
           rsCliente.Update
        End If
    End If
End If
LcSql = "UPDATE alid050 set status='CANCELADA', valorproduto=0, valornota=0 where numnf='" & NF.Text & "'"
ExecutaSql LcSql


'==> Exclui a Nota da Comissao
LcSql = "Delete From aliD201 where nf='" & NF.Text & "'"
db.Execute LcSql
Do Until RsItens.EOF
    LcCodigoItem = RsItens("Codigo")
    Cestoque.CodProduto = RsItens("codprod")
    Set RsHistorico = AbreRecordset("Select * from Historicoproduto where CodUnid='" & LcCodigoItem & "' and tipo='S' and nf='" & NF.Text & "'", True)
    If Not RsHistorico.EOF Then
       Me.Caption = "Estornando produto:" & RsItens("descricao")
       
       If IsNumeric(RsHistorico("santa")) Then LcSanta = RsHistorico("santa") Else LcSanta = 0
       If IsNumeric(RsHistorico("santa2")) Then LcSanta1 = RsHistorico("santa2") Else LcSanta1 = 0
       If IsNumeric(RsHistorico("california")) Then LcCalifornia = RsHistorico("california") Else LcCalifornia = 0
       
       Cestoque.AcrescentaEstoque 0, LcSanta, LcSanta1, LcCalifornia, , , True
       Me.Caption = "Gerando a Ficha de Estoque..."
       LcSq = "insert into HistoricoProduto (produto,descricao,santa,santa2,california,nf,data,tipo,unidade,codunid,ClienteForn) values ('"
       LcSq = LcSq & RsItens("codprod") & "','" & Cestoque.RetiraCaracter(RsItens("descricao")) & "'," & LcSanta & "," & LcSanta1 & "," & LcCalifornia
       LcSq = LcSq & ",'" & NF.Text & "','" & Format(Date, "yyyy-mm-dd") & "','CS','" & RsItens("unimed") & "','0','" & Cestoque.RetiraCaracter(RsHistorico("clienteForn")) & "')"
       LcAfetados = ExecutaSql(LcSq)
   
    End If
    RsItens.MoveNext
Loop

conexaoAdo.CommitTrans
MsgBox "Cancelamento efetuado com sucesso.", 64, "aviso"
Saida:
Me.Caption = LcCap
Exit Sub
ErrCancelamento:
conexaoAdo.RollbackTrans
MsgBox "Foi encontrado o Segunte erro cancelando a Nota Fiscal:" & Chr(13) & err.Description & "chr(13) & o cancelamento não foi realizado.", 64, "Erro Nº" & err.Number
GoTo Saida
End Sub
Sub CancelaAL()
On Error GoTo ErrCancelamento
Dim db As Database
Dim LcSql As String
'Dim RsNota As ADODB.Recordset
Dim RsItens As ADODB.Recordset
Dim RsHistorico As ADODB.Recordset
Dim RsNota As ADODB.Recordset
Dim rsCliente As Recordset

Dim Cestoque As New ControleDb
Dim LcSanta As Double
Dim LcSanta1 As Double
Dim LcCalifornia As Double
NF.Text = UCase(NF.Text)
If Len(NF.Text) = 0 Then
   MsgBox "Digite o Número da AL.", 64, "Aviso"
   NF.SetFocus
   Exit Sub
End If
'NF.Text = Right("000000" & NF.Text, 6)
conexaoAdo.BeginTrans
Set db = OpenDatabase(GLBase)
LcCap = Me.Caption
Me.Caption = "Aguarde, Desfazendo Lançamentos..."
Set RsNota = AbreRecordset("Select * from saidas where numnf='" & NF.Text & "'", True)

If RsNota.EOF Then
   MsgBox "A AL não foi encontrada.", 64, "Aviso"
   conexaoAdo.CommitTrans
   GoTo Saida
End If
If RsNota("Status") = "CANCELADA" Then
   MsgBox "Esta AL já foi cancelada.", 64, "Aviso"
   conexaoAdo.CommitTrans
   GoTo Saida
End If
Set RsItens = AbreRecordset("Select * from saidasdados where numnf='" & NF.Text & "'", True)
Me.Caption = "Marcando Nota Fiscal como Cancelada."
Me.Caption = "Excluido comissões..."
LcSql = "Delete from alid201 where nf='" & NF.Text & "'"
db.Execute LcSql
Me.Caption = "Acertando dados do Sintegra..."
LcSql = "UPDATE sintegra_50 Set situacao='S' where nf='" & NF.Text & "' and cfop='" & Replace(RsNota!CFOP, ".", "") & "'"
afetados = ExecutaSql(LcSql)

Me.Caption = "Excluido Contas a Receber do Cliente..."
LcSql = "Delete from alid015 where nf like '" & NF.Text & "%'"
ExecutaSql LcSql
'==> Estorna Credito Cliente

Set RsNota = AbreRecordset("Select * from saidas where numnf='" & NF.Text & "'", True)
If Not RsNota.EOF Then
   Set rsCliente = db.OpenRecordset("Select * from Alid001 where codigo='" & Right("00000" & RsNota!Cliente, 5) & "'")
    If UCase(RsNota!Natureza) <> UCase("VV") And UCase(RsNota!Natureza) <> "DE" And UCase(RsNota!Natureza) <> "TR" And UCase(RsNota!Natureza) <> "RE" Then
        'LcCriterioPes = "codigo='" & txt(8).Text & "'"
        'RsCliente.FindFirst LcCriterioPes
        If Not rsCliente.EOF Then
           rsCliente.Edit
         '  rsCliente("ULTCOMPRA") = CDate(txt(12).Text)
           'If Natureza.Text <> "DEVOLUCAO" And Natureza.Text <> "TRANSFERENCIA" Then
              rsCliente("CreditoUtilizado") = rsCliente("CreditoUtilizado") - CCur(RsNota!ValorNota)
           'End If
           rsCliente.Update
        End If
    End If
End If
LcSql = "UPDATE saidas set status='CANCELADA', valorproduto=0, valornota=0 where numnf='" & NF.Text & "'"
ExecutaSql LcSql


'==> Exclui a Nota da Comissao
LcSql = "Delete From aliD201 where nf='" & NF.Text & "'"
db.Execute LcSql
Do Until RsItens.EOF
    LcCodigoItem = RsItens("Codigo")
    Cestoque.CodProduto = RsItens("codprod")
    Set RsHistorico = AbreRecordset("Select * from Historicoproduto where CodUnid='" & LcCodigoItem & "' and tipo='S' and nf='" & NF.Text & "'", True)
  
    'MsgBox DEscricaoErro
    
    If Not RsHistorico.EOF Then
       Me.Caption = "Estornando produto:" & RsItens("descricao")
       
       If IsNumeric(RsHistorico("santa")) Then LcSanta = RsHistorico("santa") Else LcSanta = 0
       If IsNumeric(RsHistorico("santa2")) Then LcSanta1 = RsHistorico("santa2") Else LcSanta1 = 0
       If IsNumeric(RsHistorico("california")) Then LcCalifornia = RsHistorico("california") Else LcCalifornia = 0
       
       Cestoque.AcrescentaEstoque 0, LcSanta, LcSanta1, LcCalifornia, , , True
       Me.Caption = "Gerando a Ficha de Estoque..."
       LcSq = "insert into HistoricoProduto (produto,descricao,santa,santa2,california,nf,data,tipo,unidade,codunid,ClienteForn) values ('"
       LcSq = LcSq & RsItens("codprod") & "','" & Cestoque.RetiraCaracter(RsItens("descricao")) & "'," & LcSanta & "," & LcSanta1 & "," & LcCalifornia
       LcSq = LcSq & ",'" & NF.Text & "','" & Format(Date, "yyyy-mm-dd") & "','CS','" & RsItens("unimed") & "','0','" & Cestoque.RetiraCaracter(RsHistorico("clienteForn")) & "')"
       LcAfetados = ExecutaSql(LcSq)
   
    End If
    RsItens.MoveNext
Loop

conexaoAdo.CommitTrans
MsgBox "Cancelamento efetuado com sucesso.", 64, "aviso"
Saida:
Me.Caption = LcCap
Exit Sub
ErrCancelamento:
conexaoAdo.RollbackTrans
MsgBox "Foi encontrado o Segunte erro cancelando a Nota Fiscal:" & Chr(13) & err.Description & "chr(13) & o cancelamento não foi realizado.", 64, "Erro Nº" & err.Number
GoTo Saida
End Sub
Private Sub Command2_Click()
On Error Resume Next
Unload Me
End Sub
