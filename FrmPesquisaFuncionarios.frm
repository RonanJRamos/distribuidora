VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmPesquisaFuncionarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pesquisa Funcionarios"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11430
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   11430
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok F2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9600
      TabIndex        =   4
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar F10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9600
      TabIndex        =   5
      Top             =   720
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   2895
      Begin VB.OptionButton Parte 
         Caption         =   "Qualquer Parte"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   720
         Width           =   1815
      End
      Begin VB.OptionButton Inicio 
         Caption         =   "Inicio"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MostraCliente 
      Height          =   4095
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   7223
      _Version        =   393216
      Rows            =   17
      Cols            =   3
      BackColor       =   -2147483624
   End
   Begin VB.TextBox txt 
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   720
      Width           =   6015
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Critério"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3240
      TabIndex        =   7
      Top             =   360
      Width           =   1020
   End
End
Attribute VB_Name = "FrmPesquisaFuncionarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private LcCorAnterior, a As Integer
Private Sub CmdCancelar_Click()
On Error Resume Next
Me.Visible = False
FrmLocacao.Txt(4).SetFocus
 
End Sub

Private Sub CmdCancelar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{C}"
End Sub

Private Sub CmdOk_Click()
ExibePesquisa
End Sub

Private Sub CmdOk_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{C}"
End Sub

Private Sub Form_Activate()
On Error Resume Next
Txt.Text = ""
ExibePesquisa
End Sub

Private Sub Form_Load()
On Error Resume Next
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
GeraGrid
ExibePesquisa
End Sub
Function GeraGrid()
On Error Resume Next
MostraCliente.ColAlignment(0) = 7
MostraCliente.ColAlignment(1) = 1


MostraCliente.ColWidth(0) = 700
MostraCliente.ColWidth(1) = 5900
MostraCliente.ColWidth(2) = 0

MostraCliente.TextMatrix(0, 0) = "Código"
MostraCliente.TextMatrix(0, 1) = "Nome"

LcTamanhoGrid = 1
End Function
Function ExibePesquisa()
On Error Resume Next
Dim RsVendedor As Recordset, RsVendedor1 As Recordset, RsVendedor2 As Recordset
Dim LcCriSql, LcCriSql1, LcCriSql2 As String
Dim LcTamanho, a As Long
'Verifica se Selecionou todos
If Len(Trim(GlCriterioSql)) > 0 Then
    LcCriSql = GlCriterioSql
Else
    If Len(Trim(Txt.Text)) = 0 Then
       LcCriSql = "select * From alid200 where nome like '*' order by nome"
      msg = "Aguarde, Criando Lista de Funcionarios ..."
    Else
      msg = "Aguarde, Filtrando Funcionarios Começados com " & UCase(Txt.Text)
      If Inicio Then
        LcCriSql = "select * From alid200 where nome like '" & UCase(Txt.Text) & "*' order by nome"
      Else
        LcCriSql = "select * From alid200 where nome like '*" & UCase(Txt.Text) & "*'  order by nome"
      End If
    End If
End If
'Set DbBase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
AbreBase
Set RsVendedor = Dbbase.OpenRecordset(LcCriSql, dbOpenDynaset, dbSeeChanges, dbOptimistic) ', dbOpenDynaset)

LcTamanho = MostraCliente.Rows
a = 2
Me.Caption = msg
MostraCliente.Rows = 1

Do Until RsVendedor.EOF
  If err.Number > 0 Then Exit Do
  If Len(Trim(RsVendedor!nome)) > 0 Then
   If Not IsNull(RsVendedor!nome) Then
     
     MostraCliente.Rows = a
     MostraCliente.TextMatrix(a - 1, 0) = RsVendedor!codigo
     MostraCliente.TextMatrix(a - 1, 1) = RsVendedor!nome
     MostraCliente.TextMatrix(a - 1, 2) = RsVendedor!Comissao
     a = a + 1
     RsVendedor.MoveNext
    End If
  End If
Loop
GlCriterioSql = ""
Me.Caption = "Funcionarios Começados com " & Txt.Text

RsVendedor.Close
Set RsVendedor = Nothing
MostraCliente.SetFocus
End Function

Private Sub Inicio_Click()
Txt.SetFocus
End Sub


Private Sub Inicio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Unload Me
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{C}"
End Sub

Private Sub MostraCliente_DblClick()
 On Error Resume Next
 Dim a As Integer
 a = MostraCliente.Row

  Select Case GlFormA.Name
     Case Is = "FrmSaidaProduto"
        FrmSaidaProduto.Txt(10).Text = MostraCliente.TextMatrix(a, 0)
        FrmSaidaProduto.Txt(7).Text = MostraCliente.TextMatrix(a, 1)
        Me.Visible = False
        FrmSaidaProduto.Txt(7).SetFocus
    Case Is = "Despesas"
        Despesas.Txt(2).SetFocus
        Despesas.Txt(2).Text = MostraCliente.TextMatrix(a, 0)
        Despesas.Txt(3).Text = MostraCliente.TextMatrix(a, 1)
        Me.Visible = False
    Case Is = "Orcamento"
        orcamento.CodigoCliente.SetFocus
        orcamento.codigoVendedor.Text = MostraCliente.TextMatrix(a, 0)
        orcamento.NomeVendedor.Text = MostraCliente.TextMatrix(a, 1)
        orcamento.Comissao.Text = MostraCliente.TextMatrix(a, 2)
        orcamento.ComissaoProduto.Text = MostraCliente.TextMatrix(a, 2)
        Me.Visible = False
   Case Is = "FrmProposta"
        FrmProposta.Txt(10).Text = MostraCliente.TextMatrix(a, 0)
        FrmProposta.Txt(7).Text = MostraCliente.TextMatrix(a, 1)
        Me.Visible = False
        FrmProposta.Txt(7).SetFocus
     
  End Select
  MostraCliente.BackColor = &H80000018
  LcPerguntaVendedor = True
  GlFormA.SetFocus
End Sub

Private Sub MostraCliente_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Dim a As Integer
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{C}"
If KeyCode = 13 Then
   
   a = MostraCliente.Row
  Select Case GlFormA.Name
     Case Is = "FrmSaidaProduto"
        FrmSaidaProduto.Txt(10).Text = MostraCliente.TextMatrix(a, 0)
        FrmSaidaProduto.Txt(7).Text = MostraCliente.TextMatrix(a, 1)
        Me.Visible = False
        FrmSaidaProduto.Txt(7).SetFocus
    Case Is = "Despesas"
        Despesas.Txt(2).SetFocus
        Despesas.Txt(2).Text = MostraCliente.TextMatrix(a, 0)
        Despesas.Txt(3).Text = MostraCliente.TextMatrix(a, 1)
        Me.Visible = False
    Case Is = "Orcamento"
        orcamento.CodigoCliente.SetFocus
        orcamento.codigoVendedor.Text = MostraCliente.TextMatrix(a, 0)
        orcamento.NomeVendedor.Text = MostraCliente.TextMatrix(a, 1)
        orcamento.Comissao.Text = MostraCliente.TextMatrix(a, 2)
        orcamento.ComissaoProduto.Text = MostraCliente.TextMatrix(a, 2)
        Me.Visible = False
        FrmProposta.Txt(10).Text = MostraCliente.TextMatrix(a, 0)
        FrmProposta.Txt(7).Text = MostraCliente.TextMatrix(a, 1)
        Me.Visible = False
        FrmProposta.Txt(7).SetFocus
        
  End Select
  MostraCliente.BackColor = &H80000018
  LcPerguntaVendedor = True
  GlFormA.SetFocus
End If

End Sub

Private Sub Parte_Click()
Txt.SetFocus
End Sub

Private Sub Parte_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Unload Me
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{C}"
End Sub

Private Sub txt_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then
   MostraCliente.SetFocus
End If
If KeyCode = 27 Then Unload Me
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{C}"
End Sub


