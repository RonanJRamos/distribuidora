VERSION 5.00
Begin VB.Form CancelaOrcamneto 
   Caption         =   "Cancelamento de Orçamentos"
   ClientHeight    =   1575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   1575
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar F10"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Confirma F3"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox txt 
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   2415
   End
   Begin VB.Line Line1 
      X1              =   3000
      X2              =   3000
      Y1              =   0
      Y2              =   1560
   End
   Begin VB.Label Label1 
      Caption         =   "Nº do Orçamento"
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
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "CancelaOrcamneto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private a As Integer
Private Sub Command1_Click()
Dim RsContasReceber As Recordset
Dim RsCaixa As Recordset
Dim Rsorcamento As Recordset
Dim RsItens As Recordset
Dim RsCliente As Recordset
Dim RsProduto As Recordset
Dim RsComissao As Recordset
Dim RsComissaoR As Recordset
LcCap = Me.Caption
Me.Caption = "Aguarde, efetuando Lançamento..."
LcSql1 = "Select * from orcamento where doc='" & Right("000000" & txt.Text, 6) & "'"
LcSql2 = "Select * from DadosOrcamento where doc='" & Right("000000" & txt.Text, 6) & "'"
LcSql3 = "Select * from Alid001"
LcSql4 = "Select * from Alid009"
LcSql5 = "Select * from Alid201 where nf='" & Right("000000" & txt.Text, 6) & "'"
LcSql6 = "Select * from Alid015 where nf like '" & Right("000000" & txt.Text, 6) & "*'"
LcSql7 = "Select * from Alid016 where nf='" & Right("000000" & txt.Text, 6) & "'"
LcSql8 = "Select * from ComissaoRepresentante where nf='" & Right("000000" & txt.Text, 6) & "'"

AbreBase
Set RsContasReceber = Dbbase.OpenRecordset(LcSql6)
Set RsCaixa = Dbbase.OpenRecordset(LcSql7)
Set Rsorcamento = Dbbase.OpenRecordset(LcSql1)
Set RsItens = Dbbase.OpenRecordset(LcSql2)
Set RsCliente = Dbbase.OpenRecordset(LcSql3)
Set RsProduto = Dbbase.OpenRecordset(LcSql4)
Set RsComissao = Dbbase.OpenRecordset(LcSql5)
Set RsComissaoR = Dbbase.OpenRecordset(LcSql8)
If Rsorcamento.EOF Then
   MsgBox "O Orçamento/Vendas Nº " & Right("000000" & txt.Text, 6) & " Não foi Encontrado...", 64, "Aviso"
   Exit Sub
Else
   LcCliente = Rsorcamento!Cliente
   Rsorcamento.Edit
   Rsorcamento("Status") = "Cancelado"
   Rsorcamento.Update
End If
'=== Baixa Debito do Cliente
If Rsorcamento!condpag = "A Prazo" Then
   LcCri = "CODIGO='" & LcCliente & "'"
   RsCliente.FindFirst LcCri
   If Not RsCliente.NoMatch Then
      RsCliente.Edit
      RsCliente!CreditoUtilizado = RsCliente!CreditoUtilizado - Rsorcamento!TotalGeral
      RsCliente.Update
   End If
End If
'===Baixa Comissao
Do Until RsComissao.EOF
   RsComissao.Delete
   RsComissao.MoveNext
Loop
'=== Baixa Comissao Representada
Do Until RsComissaoR.EOF
    RsComissaoR.Delete
    RsComissaoR.MoveNext
Loop
'=== Baixa Contas a Receber
Do Until RsContasReceber.EOF
   RsContasReceber.Delete
   RsContasReceber.MoveNext
Loop
'=== Baixa o Caixa
Do Until RsCaixa.EOF
   RsCaixa.Delete
   RsCaixa.MoveNext
Loop
'== Atualiza o Estoque
If Rsorcamento!Natureza = "Ve" Then
   Do Until RsItens.EOF
      LcCrip = "COD='" & RsItens!codigoproduto & "'"
      RsProduto.FindFirst LcCrip
      Call Ficha(Right("000000" & txt.Text, 6), RsItens!codigoproduto, RsProduto!nome, RsItens!quant, RsItens!Unit, RsItens!quant * RsItens!Unit, "CS", RsCliente!Razaosoc, RsItens!unid, "")
      If Not RsProduto.NoMatch Then
         RsProduto.Edit
         RsProduto("QuantEstoque") = RsProduto("QuantEstoque") + quant
         RsProduto.Update
       End If
      RsItens.MoveNext
   Loop
End If
RsContasReceber.Close
RsCaixa.Close
Rsorcamento.Close
RsItens.Close
RsCliente.Close
RsProduto.Close
RsComissao.Close

Set RsContasReceber = Nothing
Set RsCaixa = Nothing
Set Rsorcamento = Nothing
Set RsItens = Nothing
Set RsCliente = Nothing
Set RsProduto = Nothing
Set RsComissao = Nothing
Me.Caption = LcCap
MsgBox "O Orçamento Nº: " & txt.Text & " Foi Cancelado...", 64, "Aviso"
txt.Text = ""

End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Command2_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
FrmPrincipal.SetFocus
End Sub

Private Sub txt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 121 Then SendKeys "%{F}"
If KeyCode = 114 Then SendKeys "%{C}"
End Sub
