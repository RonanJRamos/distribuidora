VERSION 5.00
Begin VB.Form FrmExclusaodeVales 
   BackColor       =   &H00D3D7C8&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exclusão de Vales"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3030
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   3030
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar"
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Excluir"
      Enabled         =   0   'False
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Vale 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº do Vale"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "FrmExclusaodeVales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private LcSq As String
Private Sub Command1_Click()
On errror GoTo errExc
Dim LcSql As String
Dim Estoque As ControleDb
Dim RsH As ADODB.Recordset
Dim RsV As ADODB.Recordset
Dim RsItens As ADODB.Recordset
Dim LcSanta As Double
Dim LcSanta1 As Double
Dim LcCalifornia As Double

If Len(Vale.Text) = 0 Then
   MsgBox "Informe o numero do Vale a ser Excluido.", 64, "Aviso"
   Vale.SetFocus
   Exit Sub
End If

LcResp = MsgBox("Confirma a Exclusão do Vale?", vbExclamation + vbYesNo, "Aviso")
If LcResp = vbNo Then Exit Sub
Vale.Text = Right("000000" & Vale.Text, 6)

conexaoAdo.BeginTrans
Set RsV = AbreRecordset("Select * from vales where NUMNF='" & Vale.Text & "'", True)

If RsV.EOF Then
   MsgBox "O vale " & Vale.Text & " Não foi encontrado.", vbOKOnly, "Aviso"
   conexaoAdo.RollbackTrans
   GoTo Saida
End If

'===> Efetua o estorno
Set RsItens = AbreRecordset("Select * from valesProdutos where NumNf='" & Vale.Text & "'", True)

Set Estoque = New ControleDb
Do Until RsItens.EOF
    LcCodigoItem = RsItens("Codigo")
    Set RsH = AbreRecordset("Select * from HistoricoProduto where CodUnid='" & LcCodigoItem & "' and tipo='V'", True)

    Estoque.CodProduto = RsItens("CodProd")
    Estoque.Nf = Vale.Text
    'If Not Estoque.EstornaEstoque(0, CDbl(RsH!santa), CDbl(RsH!Santa2), CDbl(RsH!California), 0, RsH!Unidade) Then
    '   err.Raise vbObjectError + 513, "ErroEstornando o Estoque.", "Erro Estornando o Estoque"
    '   GoTo errExc
    'End If
    If IsNumeric(RsH("santa")) Then LcSanta = RsH("santa") Else LcSanta = 0
    If IsNumeric(RsH("santa2")) Then LcSanta1 = RsH("santa2") Else LcSanta1 = 0
    If IsNumeric(RsH("california")) Then LcCalifornia = RsH("california") Else LcCalifornia = 0

    Estoque.AcrescentaEstoque 0, LcSanta, LcSanta1, LcCalifornia, , , True

    LcQSanta = RsH!santa
    LcQSanta1 = RsH!Santa2
    LcqCanifornia = RsH!California
    LcUn = RsH!Unidade & ""
    LcCliente = RsH!clienteforn & ""
    LcSq = "insert into HistoricoProduto (produto,descricao,santa,santa2,california,nf,data,tipo,unidade,codunid,ClienteForn) values ('"
    LcSq = LcSq & Estoque.CodProduto & "','" & Estoque.RetiraCaracter(Estoque.DescricaoProduto) & "'," & LcQSanta & "," & LcQSanta1 & "," & LcqCanifornia
    LcSq = LcSq & ",'" & Estoque.Nf & "','" & Format(Date, "yyyy-mm-dd") & "','EV','" & LcUn & "','0','" & LcCliente & "')"
    
    ExecutaSql LcSq

    RsItens.MoveNext
Loop
LcSql = "Delete from vales where NUMNF='" & Vale.Text & "'"
ExecutaSql LcSql
LcSql = "Delete from valesprodutos where NUMNF='" & Vale.Text & "'"
ExecutaSql LcSql
conexaoAdo.CommitTrans
MsgBox "Exclusão Efetuada com sucesso.", 64, "Aviso"
Saida:
On Error Resume Next
Set RsV = Nothing
Set RsH = Nothing
Set Estoque = Nothing

Exit Sub
errExc:
conexaoAdo.RollbackTrans
MsgBox "Ocorreu o seguinte erro Excluido o vale:" & Chr(13) & err.Description & Chr(13) & Chr(13) & "A Exclusão foi cancelada.", 64, "Aviso"
GoTo Saida

End Sub

Private Sub Command2_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub Vale_Change()
On Error Resume Next
Command1.Enabled = Len(Vale.Text)

End Sub
