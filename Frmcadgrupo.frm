VERSION 5.00
Begin VB.Form FrmCadGrupo 
   BackColor       =   &H00B3E9FD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Grupos de Acesso"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7305
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   7305
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdExcluir 
      Caption         =   "&Excluir F3"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   14
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00B3E9FD&
      Caption         =   "Check1"
      Height          =   495
      Index           =   4
      Left            =   4320
      TabIndex        =   13
      Top             =   3480
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00B3E9FD&
      Caption         =   "Check1"
      Height          =   495
      Index           =   3
      Left            =   4320
      TabIndex        =   10
      Top             =   3040
      Width           =   2295
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00B3E9FD&
      Caption         =   "Check1"
      Height          =   495
      Index           =   2
      Left            =   4320
      TabIndex        =   9
      Top             =   2600
      Width           =   2295
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00B3E9FD&
      Caption         =   "Check1"
      Height          =   495
      Index           =   1
      Left            =   4320
      TabIndex        =   8
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00B3E9FD&
      Caption         =   "Acesso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   4080
      TabIndex        =   6
      Top             =   1320
      Width           =   2775
      Begin VB.CheckBox Check1 
         BackColor       =   &H00B3E9FD&
         Caption         =   "Check1"
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.CommandButton CmdFechar 
      Caption         =   "&Fechar F10"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton CmdSalvar 
      Caption         =   "&Salvar F2"
      Enabled         =   0   'False
      Height          =   615
      Left            =   4080
      TabIndex        =   4
      Top             =   4320
      Width           =   2775
   End
   Begin VB.CommandButton CmdNovo 
      Caption         =   "&Novo Grupo F4"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   4560
      Width           =   1335
   End
   Begin VB.ListBox Lista 
      Height          =   2790
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   2655
   End
   Begin VB.ComboBox CmbRecurso 
      Height          =   315
      Left            =   4200
      TabIndex        =   1
      Top             =   600
      Width           =   2655
   End
   Begin VB.TextBox txt 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Função do Sistema"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   4200
      TabIndex        =   12
      Top             =   120
      Width           =   2340
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grupo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   480
      TabIndex        =   11
      Top             =   240
      Width           =   765
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000001&
      X1              =   0
      X2              =   3960
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000001&
      X1              =   3960
      X2              =   0
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      X1              =   3960
      X2              =   3960
      Y1              =   0
      Y2              =   5040
   End
End
Attribute VB_Name = "FrmCadGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type tipoM
        Funcao As String
        CodigoGrupo As Long
        DescricaoGrupo As String
        Permissao As Long
        Incluir As Integer
        Alterar As Integer
        Consultar As Integer
        Relatorio As Integer
        Baixa As Integer
End Type
Private MtGrupo() As tipoM
Private LcTamanhoMa, a As Long
Private LcNovo, LcSalvo As Integer

Private Sub Check1_Click(Index As Integer)
On Error Resume Next
If Len(Trim(Txt.Text)) = 0 Then
   MsgBox "Digite ou Escolha o Grupo a Alterar", 48, "Aviso"
   CmbRecurso.Text = ""
   Txt.SetFocus
Else
  CmdSalvar.Enabled = True
End If
End Sub

Private Sub Check1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 113 Then SendKeys "%+{S}"
If KeyCode = 114 Then SendKeys "{E}"
If KeyCode = 115 Then SendKeys "{N}"
End Sub

Private Sub CmbRecurso_Change()
On Error Resume Next
If Len(Trim(Txt.Text)) = 0 Then
   MsgBox "Digite ou Escolha o Grupo a Alterar", 48, "Aviso"
   CmbRecurso.Text = ""
   Txt.SetFocus
End If
End Sub

Private Sub CmbRecurso_Click()
If LcSalvo Then
   lcrespsta = MsgBox("As Atribuições Atuais Não Foram Salvas, Salvas Agora ?", 36, "Aviso")
   If Resposta = 6 Then
      GravaPermissoesAtuais
      LcSalvo = False
   End If
   LcSalvo = False
End If
LcSalvo = False
If Len(Trim(Txt.Text)) = 0 Then
   MsgBox "Digite ou Escolha o Grupo a Alterar", 48, "Aviso"
   CmbRecurso.Text = ""
   Txt.SetFocus
Else
   
   MontaDescricaofuncao (CmbRecurso.Text)
   Call ExibePermissao(CmbRecurso.ListIndex)
End If
LcSalvo = False
CmdSalvar.Enabled = False
End Sub

Private Sub CmbRecurso_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then SendKeys "{TAB}"

If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 113 Then SendKeys "%+{S}"
If KeyCode = 114 Then SendKeys "{E}"
If KeyCode = 115 Then SendKeys "{N}"
End Sub

Private Sub CmdExcluir_Click()
On Error GoTo erroexcluir

Dim RsGrupo As ADODB.Recordset, RsUsuario As ADODB.Recordset
Dim a, Item, LcResposta As Long

Dim LcCriterio As String, LcCriterio1 As String
LcCriterio = "Select * From GrpSenhas where Grupo='" & Txt.Text & "'"
LcCriterio1 = "Select * From usuario where grupo='" & Txt.Text & "'"

'Set Dbbase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
Set RsGrupo = AbreRecordset(LcCriterio)
Set RsUsuario = AbreRecordset(LcCriterio1)
a = 0
LcResposta = MsgBox("A Exclusão do Grupo " & Txt.Text & Chr(13) & " Causará a Exclusão dos Usuários Cadastrados Neste Grupo." & Chr(13) & "Confirma a Exclusão ?", 36, "Aviso")
If LcResposta = 7 Then
   Exit Sub
End If
Do Until RsGrupo.EOF
   RsGrupo.Delete
   RsGrupo.MoveNext
Loop

Do Until RsUsuario.EOF
   RsUsuario.Delete
   RsUsuario.MoveNext
Loop
For Item = Lista.ListCount - 1 To 0 Step -1
    If Lista.List(Item) = Txt.Text Then
       Lista.RemoveItem (Item)
       Exit For
    End If
Next

RsUsuario.Close
RsGrupo.Close
Set RsUsuario = Nothing
Set RsGrupo = Nothing

Check1(0).Visible = False
Check1(1).Visible = False
Check1(2).Visible = False
Check1(3).Visible = False
Check1(4).Visible = False
CmbRecurso.Text = ""
Exit Sub
erroexcluir:

Select Case ErrosSistema
       Case Is = 0
          Resume Next
       Case Is = 6
          Resume 0
       Case Is = 7
          End
       Case Else
         If err = 13 Then
             MsgBox "Preenchimento de campo Incorreto. Favor Verificar.", 48, "Aviso"
             Exit Sub
         Else
             MsgBox err.Description & " N " & err
             Resume Next
         End If
End Select
End Sub

Private Sub CmdExcluir_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 113 Then SendKeys "%+{S}"
If KeyCode = 114 Then SendKeys "{E}"
If KeyCode = 115 Then SendKeys "{N}"
End Sub

Private Sub CmdFechar_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub CmdFechar_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 113 Then SendKeys "%+{S}"
If KeyCode = 114 Then SendKeys "{E}"
If KeyCode = 115 Then SendKeys "{N}"
End Sub

Private Sub CmdNovo_Click()
LcNovo = True
Check1(0).Visible = False
Check1(1).Visible = False
Check1(2).Visible = False
Check1(3).Visible = False
Check1(4).Visible = False
Txt.Text = ""
CmbRecurso.Text = ""
Txt.SetFocus
End Sub

Private Sub CmdNovo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 113 Then SendKeys "%+{S}"
If KeyCode = 114 Then SendKeys "{E}"
If KeyCode = 115 Then SendKeys "{N}"
End Sub

Private Sub CmdSalvar_Click()
GravaPermissoesAtuais
End Sub

Private Sub CmdSalvar_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 113 Then SendKeys "%+{S}"
If KeyCode = 114 Then SendKeys "{E}"
If KeyCode = 115 Then SendKeys "{N}"
End Sub

Private Sub Form_Activate()
Set GlFormA = Me
End Sub

Private Sub Form_Load()
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
CarregaCombo
CarregaListaInicio
Check1(0).Visible = False
Check1(1).Visible = False
Check1(2).Visible = False
Check1(3).Visible = False
Check1(4).Visible = False
End Sub
Function CarregaCombo()


CmbRecurso.AddItem "Clientes"
CmbRecurso.AddItem "Fornecedores"
CmbRecurso.AddItem "Funcionarios"
CmbRecurso.AddItem "Produtos"
CmbRecurso.AddItem "Tipo Receitas e Despesas"
CmbRecurso.AddItem "Galpão"
CmbRecurso.AddItem "Cidades"
CmbRecurso.AddItem "Tipo Monetário"
CmbRecurso.AddItem "Unidade"
CmbRecurso.AddItem "Transportadora"
CmbRecurso.AddItem "Custo"
CmbRecurso.AddItem "Receitas"
CmbRecurso.AddItem "Despesas"
CmbRecurso.AddItem "Cheques"
CmbRecurso.AddItem "Comissões"
CmbRecurso.AddItem "Comissões Representada"
CmbRecurso.AddItem "Caixa"
CmbRecurso.AddItem "Vales"
CmbRecurso.AddItem "Excluir Vales"

CmbRecurso.AddItem "Entrada de produto"
CmbRecurso.AddItem "Solicitação de Compra"
CmbRecurso.AddItem "Orçamento e Vendas"
CmbRecurso.AddItem "Pedido de Vendas"
CmbRecurso.AddItem "Alteração de Preço"
CmbRecurso.AddItem "Pesquisa Compras de Cliente"
CmbRecurso.AddItem "Ficha de Estoque"
CmbRecurso.AddItem "Romaneio"
CmbRecurso.AddItem "Gerar disquete Receita"

CmbRecurso.AddItem "Cancelar Pedidos"
CmbRecurso.AddItem "Notas de Saídas"
CmbRecurso.AddItem "Cancelar Notas"
CmbRecurso.AddItem "Copia de Segurança"
CmbRecurso.AddItem "Localizar banco de dados"
CmbRecurso.AddItem "Ver dados Excluidos"

CmbRecurso.AddItem "Senha"
CmbRecurso.AddItem "Pano de Fundo"
CmbRecurso.AddItem "Reparar banco de Dados"
CmbRecurso.AddItem "Opções"
CmbRecurso.AddItem "Alterar Data"
CmbRecurso.AddItem "Configura Mala Direta"
End Function
Function MontaDescricaofuncao(LcNome As String)

Check1(0).Visible = True
Check1(1).Visible = True
Check1(2).Visible = True
Check1(3).Visible = True
Check1(4).Visible = True
Debug.Print LcNome
Select Case LcNome

         Case Is = "Custo"
             Check1(0).Caption = "Incluir"
             Check1(1).Caption = "Alterar"
             Check1(2).Caption = "Consultar"
             Check1(3).Visible = False
             Check1(4).Visible = False
         Case Is = "Comissões Representada"
             Check1(0).Caption = "Sim"
             Check1(1).Visible = False
             Check1(2).Visible = False
             Check1(3).Visible = False
             Check1(4).Visible = False
         Case Is = "Pedido de Vendas"
             Check1(0).Caption = "Sim"
             Check1(1).Visible = False
             Check1(2).Visible = False
             Check1(3).Visible = False
             Check1(4).Visible = False
         Case Is = "Orçamento e Vendas"
             Check1(0).Caption = "Sim"
             Check1(1).Caption = "Relatórios"
             'Check1(1).Visible = False
             Check1(2).Visible = False
             Check1(3).Visible = False
             Check1(4).Visible = False
         Case Is = "Pesquisa Compras de Cliente"
             Check1(0).Caption = "Sim"
             Check1(1).Visible = False
             Check1(2).Visible = False
             Check1(3).Visible = False
             Check1(4).Visible = False
        Case Is = "Copia de Segurança"
             Check1(0).Caption = "Backup"
             Check1(1).Caption = "Recuperar"
             Check1(2).Visible = False
             Check1(3).Visible = False
             Check1(4).Visible = False
             
        Case Is = "Clientes"
             Check1(0).Caption = "Incluir"
             Check1(1).Caption = "Alterar"
             Check1(2).Caption = "Consultar"
             Check1(3).Caption = "Relatórios"
             Check1(4).Visible = False
        Case Is = "Cidades"
             Check1(0).Caption = "Incluir"
             Check1(1).Caption = "Alterar"
             Check1(2).Caption = "Consultar"
             Check1(3).Caption = "Relatórios"
             Check1(4).Visible = False
        Case Is = "Fornecedores"
             Check1(0).Caption = "Incluir"
             Check1(1).Caption = "Alterar"
             Check1(2).Caption = "Consultar"
             Check1(3).Caption = "Relatórios"
             Check1(4).Visible = False
        Case Is = "Funcionarios"
             Check1(0).Caption = "Incluir"
             Check1(1).Caption = "Alterar"
             Check1(2).Caption = "Consultar"
             Check1(3).Caption = "Relatórios"
             Check1(4).Visible = False
        Case Is = "Produtos"
             Check1(0).Caption = "Incluir"
             Check1(1).Caption = "Alterar"
             Check1(2).Caption = "Consultar"
             Check1(3).Caption = "Relatórios"
             Check1(4).Visible = False
        Case Is = "Tipo Receitas e Despesas"
             Check1(0).Caption = "Incluir"
             Check1(1).Caption = "Alterar"
             Check1(2).Caption = "Consultar"
             Check1(3).Caption = "Relatórios"
             Check1(4).Visible = False
        Case Is = "Galpão"
             Check1(0).Caption = "Incluir"
             Check1(1).Caption = "Alterar"
             Check1(2).Caption = "Consultar"
             Check1(3).Caption = "Relatórios"
             Check1(4).Visible = False
             
        Case Is = "Transportadora"
             Check1(0).Caption = "Incluir"
             Check1(1).Caption = "Alterar"
             Check1(2).Caption = "Consultar"
             Check1(3).Caption = "Relatórios"
             Check1(4).Visible = False
        Case Is = "Tipo Monetário"
             Check1(0).Caption = "Incluir"
             Check1(1).Caption = "Alterar"
             Check1(2).Caption = "Consultar"
             Check1(3).Caption = "Relatórios"
             Check1(4).Visible = False
        Case Is = "Unidade"
             Check1(0).Caption = "Incluir"
             Check1(1).Caption = "Alterar"
             Check1(2).Caption = "Consultar"
             Check1(3).Caption = "Relatório"
             Check1(4).Visible = False
        Case Is = "Despesas"
             Check1(0).Caption = "Incluir"
             Check1(1).Caption = "Alterar"
             Check1(2).Caption = "Consultar"
             Check1(3).Caption = "Relatório"
             Check1(4).Caption = "Baixa"
             Check1(4).Visible = True
        Case Is = "Receitas"
             Check1(0).Caption = "Incluir"
             Check1(1).Caption = "Alterar"
             Check1(2).Caption = "Consultar"
             Check1(3).Caption = "Relatório"
             Check1(4).Caption = "Baixa"
             Check1(4).Visible = True
        Case Is = "Cheques"
              Check1(0).Caption = "Incluir"
             Check1(1).Caption = "Alterar"
             Check1(2).Caption = "Consultar"
             Check1(3).Caption = "Relatório"
             Check1(4).Visible = False
        Case Is = "Caixa"
             Check1(0).Caption = "Fechamento"
             Check1(1).Caption = "Relatório"
             Check1(2).Visible = False
             Check1(3).Visible = False
             Check1(4).Visible = False
             Check1(4).Visible = False
        Case Is = "Pedido de Cliente"
             Check1(0).Caption = "Sim"
             Check1(1).Caption = "Relatório"
             Check1(2).Visible = False
             Check1(3).Visible = False
             Check1(4).Visible = False
             Check1(4).Visible = False
         Case Is = "Cancelar Pedidos"
             Check1(0).Caption = "Sim"
             Check1(1).Visible = False
             Check1(2).Visible = False
             Check1(3).Visible = False
             Check1(4).Visible = False
             Check1(4).Visible = False
         Case Is = "Cancelar Notas"
             Check1(0).Caption = "Sim"
             Check1(1).Visible = False
             Check1(2).Visible = False
             Check1(3).Visible = False
             Check1(4).Visible = False
             Check1(4).Visible = False
        Case Is = "Entrada de produto"
             Check1(0).Caption = "Sim"
             Check1(1).Caption = "Relatório"
             Check1(2).Visible = False
             Check1(3).Visible = False
             Check1(4).Visible = False
             Check1(4).Visible = False
        Case Is = "Solicitação de Compra"
             Check1(0).Caption = "Sim"
             Check1(1).Caption = "Relatório"
             Check1(2).Visible = False
             Check1(3).Visible = False
             Check1(4).Visible = False
             Check1(4).Visible = False
        Case Is = "Alteração de Preço"
             Check1(0).Caption = "Sim"
             Check1(1).Visible = False
             Check1(2).Visible = False
             Check1(3).Visible = False
             Check1(4).Visible = False
             Check1(4).Visible = False
        Case Is = "Notas de Saídas"
             Check1(0).Caption = "Sim"
             Check1(1).Caption = "Relatório"
             Check1(2).Visible = False
             Check1(3).Visible = False
             Check1(4).Visible = False
             Check1(4).Visible = False
        Case Is = "Entrada de produto"
             Check1(0).Caption = "Sim"
             Check1(1).Caption = "Relatório"
             Check1(2).Visible = False
             Check1(3).Visible = False
             Check1(4).Visible = False
             Check1(4).Visible = False
        Case Is = "Solicitação de Compra"
             Check1(0).Caption = "Sim"
             Check1(1).Caption = "Relatório"
             Check1(2).Visible = False
             Check1(3).Visible = False
             Check1(4).Visible = False
             Check1(4).Visible = False
        Case Is = "Segurança"
             Check1(0).Caption = "Backup"
             Check1(1).Caption = "Respaurar"
             Check1(2).Visible = False
             Check1(3).Visible = False
             Check1(4).Visible = False
             Check1(4).Visible = False
        Case Is = "Senha"
             Check1(0).Caption = "Usuário"
             Check1(1).Caption = "Grupo"
             Check1(2).Visible = False
             Check1(3).Visible = False
             Check1(4).Visible = False
             Check1(4).Visible = False
        Case Is = "Pano de Fundo"
             Check1(0).Caption = "Sim"
             Check1(1).Visible = False
             Check1(2).Visible = False
             Check1(3).Visible = False
             Check1(4).Visible = False
             Check1(4).Visible = False
        Case Is = "Reparar banco de Dados"
             Check1(0).Caption = "Sim"
             Check1(1).Visible = False
             Check1(2).Visible = False
             Check1(3).Visible = False
             Check1(4).Visible = False
             Check1(4).Visible = False
        Case Is = "Comissões"
             Check1(0).Caption = "Sim"
             Check1(1).Visible = False
             Check1(2).Visible = False
             Check1(3).Visible = False
             Check1(4).Visible = False
             Check1(4).Visible = False
        Case Is = "Opções"
             Check1(0).Caption = "Sim"
             Check1(1).Visible = False
             Check1(2).Visible = False
             Check1(3).Visible = False
             Check1(4).Visible = False
             Check1(4).Visible = False
        Case Is = "Alterar Data"
             Check1(0).Caption = "Sim"
             Check1(1).Visible = False
             Check1(2).Visible = False
             Check1(3).Visible = False
             Check1(4).Visible = False
             Check1(4).Visible = False
       Case Is = "Configura Mala Direta"
             Check1(0).Caption = "Sim"
             Check1(1).Visible = False
             Check1(2).Visible = False
             Check1(3).Visible = False
             Check1(4).Visible = False
             Check1(4).Visible = False
        Case Is = "Fechamento do caixa"
             Check1(0).Caption = "Sim"
             Check1(1).Visible = False
             Check1(2).Visible = False
             Check1(3).Visible = False
             Check1(4).Visible = False
             Check1(4).Visible = False
        Case Is = "Vales"
             Check1(0).Caption = "Sim"
             Check1(1).Visible = False
             Check1(2).Visible = False
             Check1(3).Visible = False
             Check1(4).Visible = False
             Check1(4).Visible = False
        Case Is = "Excluir Vales"
             Check1(0).Caption = "Sim"
             Check1(1).Visible = False
             Check1(2).Visible = False
             Check1(3).Visible = False
             Check1(4).Visible = False
             Check1(4).Visible = False
        Case Is = "Ficha de Estoque"
             Check1(0).Caption = "Sim"
             Check1(1).Visible = False
             Check1(2).Visible = False
             Check1(3).Visible = False
             Check1(4).Visible = False
             Check1(4).Visible = False
        Case Is = "Romaneio"
             Check1(0).Caption = "Sim"
             Check1(1).Visible = False
             Check1(2).Visible = False
             Check1(3).Visible = False
             Check1(4).Visible = False
             Check1(4).Visible = False
        Case Is = "Gerar disquete Receita"
             Check1(0).Caption = "Sim"
             Check1(1).Visible = False
             Check1(2).Visible = False
             Check1(3).Visible = False
             Check1(4).Visible = False
             Check1(4).Visible = False
         Case Is = "Copia de Segurança"
             Check1(0).Caption = "Sim"
             Check1(1).Visible = False
             Check1(2).Visible = False
             Check1(3).Visible = False
             Check1(4).Visible = False
             Check1(4).Visible = False
     
         Case Is = "Localizar banco de dados"
             Check1(0).Caption = "Sim"
             Check1(1).Visible = False
             Check1(2).Visible = False
             Check1(3).Visible = False
             Check1(4).Visible = False
             Check1(4).Visible = False
        Case Is = "Ver dados Excluidos"
             Check1(0).Caption = "Sim"
             Check1(1).Visible = False
             Check1(2).Visible = False
             Check1(3).Visible = False
             Check1(4).Visible = False
             Check1(4).Visible = False

    End Select
End Function

Function CarregaLista()
On Error GoTo ErroCar

Dim RsGrupo As ADODB.Recordset
Dim a As Long

Dim LcCriterio As String
LcCriterio = "Select * From GrpSenhas where Grupo='" & Txt.Text & "'"
'Set Dbbase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
Set RsGrupo = AbreRecordset(LcCriterio)
a = 0
LcNovo = True

ReDim MtGrupo(0)
Do Until RsGrupo.EOF
   ReDim Preserve MtGrupo(a)
   
   MtGrupo(a).CodigoGrupo = RsGrupo!Codigo
   MtGrupo(a).DescricaoGrupo = RsGrupo!Grupo
   MtGrupo(a).Funcao = RsGrupo!Sistema
   MtGrupo(a).Incluir = RsGrupo!Incluir
   MtGrupo(a).Alterar = RsGrupo!Alterar
   MtGrupo(a).Consultar = RsGrupo!Consultar
   MtGrupo(a).Relatorio = RsGrupo!Relatorio
   MtGrupo(a).Baixa = RsGrupo!Baixa
   
   a = a + 1
   RsGrupo.MoveNext
   LcNovo = False
   
Loop
If a > 0 Then LcTamanhoMa = a - 1
RsGrupo.Close
'Dbbase.Close
Set RsGrupo = Nothing
'Set Dbbase = Nothing

Exit Function
ErroCar:

Select Case ErrosSistema
       Case Is = 0
          Resume Next
       Case Is = 6
          Resume 0
       Case Is = 7
          End
       Case Else
         If err = 13 Then
             MsgBox "Preenchimento de campo Incorreto. Favor Verificar.", 48, "Aviso"
             Exit Function
         Else
             MsgBox err.Description & " N " & err
             Resume Next
         End If
End Select
End Function

Function ExibePermissao(LcIndice As Long)
On Error Resume Next
Dim LcPermicao, a As Long

For a = 0 To UBound(MtGrupo)
    If MtGrupo(a).Funcao = CmbRecurso.Text Then
       LcPermicao = MtGrupo(a).Permissao
       Exit For
    End If
Next
If err <> 0 Then
   LcPermicao = 0
End If

If MtGrupo(a).Incluir Then Check1(0) = 1 Else Check1(0) = 0
If MtGrupo(a).Alterar Then Check1(1) = 1 Else Check1(1) = 0
If MtGrupo(a).Consultar Then Check1(2) = 1 Else Check1(2) = 0
If MtGrupo(a).Relatorio Then Check1(3) = 1 Else Check1(3) = 0
If MtGrupo(a).Baixa Then Check1(4) = 1 Else Check1(4) = 0


End Function
Function GravaRegistroNovo()
On Error GoTo errograva

Dim RsGrupo As ADODB.Recordset
Dim a, Item As Long
'Set Dbbase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
Set RsGrupo = AbreRecordset("GrpSenhas")
a = 0
LcNovo = True

For Item = CmbRecurso.ListCount - 1 To 0 Step -1
    RsGrupo.AddNew
    RsGrupo!Grupo = Txt.Text
    RsGrupo!Sistema = CmbRecurso.List(Item)
    RsGrupo!Incluir = Check1(0)
    RsGrupo!Alterar = Check1(1)
    RsGrupo!Consultar = Check1(2)
    RsGrupo!Relatorio = Check1(3)
    RsGrupo!Baixa = Check1(4)
    RsGrupo.Update
    LcSalvo = False
    LcNovo = False
Next
Lista.AddItem Txt.Text
RsGrupo.Close
Set RsGrupo = Nothing

Exit Function
errograva:
  
Select Case ErrosSistema
       Case Is = 0
          Resume Next
       Case Is = 6
          Resume 0
       Case Is = 7
          End
       Case Else
         If err = 13 Then
             MsgBox "Preenchimento de campo Incorreto. Favor Verificar.", 48, "Aviso"
             Exit Function
         Else
             MsgBox err.Description & " N " & err
             Resume Next
         End If
End Select
End Function



Private Sub Form_Unload(Cancel As Integer)
If LcSalvo Then
   lcrespsta = MsgBox("As Atribuições Atuais Não Foram Salvas, Salvas Agora ?", 36, "Aviso")
   If Resposta = 6 Then
      GravaPermissoesAtuais
      LcSalvo = False
   End If
   LcSalvo = False
End If
FrmPrincipal.SetFocus
End Sub

Private Sub Lista_DblClick()
On Error Resume Next
Txt.Text = Lista.Text
CarregaLista
Check1(0).Visible = False
Check1(1).Visible = False
Check1(2).Visible = False
Check1(3).Visible = False
Check1(4).Visible = False
CmbRecurso.Text = ""
CmdExcluir.Enabled = True
End Sub

Private Sub Lista_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"

If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 113 Then SendKeys "%+{S}"
If KeyCode = 114 Then SendKeys "{E}"
If KeyCode = 115 Then SendKeys "{N}"
End Sub

Private Sub Txt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"

If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 113 Then SendKeys "%+{S}"
If KeyCode = 114 Then SendKeys "{E}"
If KeyCode = 115 Then SendKeys "{N}"
End Sub

Private Sub Txt_LostFocus()

If LcNovo Then
   If Len(Trim(Txt.Text)) <> 0 Then
      GravaRegistroNovo
      CarregaLista
   End If
Else
   
   If Len(Trim(Txt.Text)) <> 0 Then
    If Not VerificaLista(Txt.Text) Then
       MsgBox "Digite um Item Já cadastrado, Ou escolha na lista Abaixo..." & Chr(13) & "Para Um Novo Item Clique no Botão Novo", 48, "Aviso"
       Txt.SetFocus
    Else
      CarregaLista
    End If
  End If
End If

End Sub
Function CarregaListaInicio()
On Error GoTo errolista

Dim RsGrupo As ADODB.Recordset
Dim a As Long

Dim LcCriterio As String
'LcCriterio = "Select * From GrpSenhas where nome='" & txt.Text & "'"
'Set Dbbase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
Set RsGrupo = AbreRecordset("select * from GrpSenhas")
a = 0
LcNovo = True
Lista.Clear
Do Until RsGrupo.EOF
   ReDim Preserve MtGrupo(a)
   If Not VerificaLista(RsGrupo!Grupo) Then Lista.AddItem RsGrupo!Grupo
   RsGrupo.MoveNext
   LcNovo = False
   
Loop

If a > 0 Then LcTamanhoMa = a - 1
RsGrupo.Close
Set RsGrupo = Nothing
Exit Function
errolista:

Select Case ErrosSistema
       Case Is = 0
          Resume Next
       Case Is = 6
          Resume 0
       Case Is = 7
          End
       Case Else
         If err = 13 Then
             MsgBox "Preenchimento de campo Incorreto. Favor Verificar.", 48, "Aviso"
             Exit Function
         Else
             MsgBox err.Description & " N " & err
             Resume Next
         End If
End Select
End Function
Function VerificaLista(LcNome As String) As Integer
Dim Item As Long

For Item = Lista.ListCount - 1 To 0 Step -1
     If Lista.List(Item) = LcNome Then
        VerificaLista = True
        Exit For
     Else
        VerificaLista = False
     End If
Next
     
    
End Function

Function GravaPermissoesAtuais()

On Error GoTo erroperm

Dim RsGrupo As ADODB.Recordset
Dim a, LcPermicao As Long
Dim LcNome As String
Dim LcCriterio As String, LcPesquisa As String

Dim LcAchou As Integer
a = 0
LcNovo = False
'Lista.Clear
'ReDim MtGrupo(0)
LcNome = CmbRecurso.Text
   
For a = 0 To UBound(MtGrupo)
    If MtGrupo(a).Funcao = LcNome Then
       
          '===>> Atualiza a matriz
          LcCriterio = "Select * From GrpSenhas where Grupo='" & Txt.Text & "' and Sistema='" & LcNome & "'"
          'Set Dbbase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
           Set RsGrupo = AbreRecordset(LcCriterio)

         
          If Not RsGrupo.EOF Then
              'RsGrupo.Edit
          Else
              RsGrupo.AddNew
              RsGrupo!Grupo = Txt.Text
              RsGrupo!Sistema = LcNome
          End If
          
          MtGrupo(a).Incluir = Check1(0)
          MtGrupo(a).Alterar = Check1(1)
          MtGrupo(a).Consultar = Check1(2)
          MtGrupo(a).Relatorio = Check1(3)
          MtGrupo(a).Baixa = Check1(4)
          
          RsGrupo!Incluir = Check1(0)
          RsGrupo!Alterar = Check1(1)
          RsGrupo!Consultar = Check1(2)
          RsGrupo!Relatorio = Check1(3)
          RsGrupo!Baixa = Check1(4)
          RsGrupo.Update
          LcAchou = True
          Exit For
          
   End If
Next
If Not LcAchou Then
           LcCriterio = "Select * From GrpSenhas where Grupo='" & Txt.Text & "' and Sistema='" & LcNome & "'"
          'Set Dbbase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
           Set RsGrupo = AbreRecordset(LcCriterio)

          RsGrupo.AddNew
          RsGrupo!Grupo = Txt.Text
          RsGrupo!Sistema = LcNome
         ' MtGrupo(a).Incluir = Check1(0)
         ' MtGrupo(a).Alterar = Check1(1)
         ' MtGrupo(a).Consultar = Check1(2)
         ' MtGrupo(a).Relatorio = Check1(3)
         ' MtGrupo(a).Baixa = Check1(4)
          RsGrupo!Incluir = Check1(0)
          RsGrupo!Alterar = Check1(1)
          RsGrupo!Consultar = Check1(2)
          RsGrupo!Relatorio = Check1(3)
          RsGrupo!Baixa = Check1(4)
          RsGrupo.Update
End If

RsGrupo.Close
Set RsGrupo = Nothing
LcNovo = False
LcSalvo = False
CmdSalvar.Enabled = False
Exit Function
erroperm:
  MsgBox err.Description & err.Number
Select Case ErrosSistema
       Case Is = 0
          Resume Next
       Case Is = 6
          Resume 0
       Case Is = 7
          End
       Case Else
         If err = 13 Then
             MsgBox "Preenchimento de campo Incorreto. Favor Verificar.", 48, "Aviso"
             Exit Function
         Else
             MsgBox err.Description & " N " & err
             Resume Next
         End If
End Select
End Function
