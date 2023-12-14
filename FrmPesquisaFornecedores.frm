VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmPesquisaFornecedores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pesquisa Fornecedores"
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
      TabIndex        =   0
      Top             =   1440
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   7223
      _Version        =   393216
      Rows            =   17
      Cols            =   5
      FixedCols       =   0
      BackColor       =   -2147483624
      FocusRect       =   2
      SelectionMode   =   1
   End
   Begin VB.TextBox txt 
      Height          =   375
      Left            =   3240
      TabIndex        =   1
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
Attribute VB_Name = "FrmPesquisaFornecedores"
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
Me.Refresh
Txt.Text = ""
ExibePesquisa
MostraCliente.SetFocus
End Sub

Private Sub Form_Load()
On Error Resume Next
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
GeraGrid
ExibePesquisa
MostraCliente.SetFocus
End Sub
Function GeraGrid()
On Error Resume Next
MostraCliente.ColAlignment(0) = 7
MostraCliente.ColAlignment(1) = 1
MostraCliente.ColAlignment(2) = 1
MostraCliente.ColAlignment(3) = 1
MostraCliente.ColAlignment(4) = 1

MostraCliente.ColWidth(0) = 700
MostraCliente.ColWidth(1) = 3900
MostraCliente.ColWidth(2) = 2400
MostraCliente.ColWidth(3) = 3900
MostraCliente.ColWidth(4) = 1000

MostraCliente.TextMatrix(0, 0) = "Código"
MostraCliente.TextMatrix(0, 1) = "Nome"
MostraCliente.TextMatrix(0, 2) = "C.N.P.J.:"
MostraCliente.TextMatrix(0, 3) = "Endereço"
MostraCliente.TextMatrix(0, 4) = "Fone"
LcTamanhoGrid = 1
End Function
Function ExibePesquisa()
On Error Resume Next
Dim RsFornecedores As Recordset, RsFornecedores1 As Recordset, RsFornecedores2 As Recordset
Dim LcCriSql, LcCriSql1, LcCriSql2 As String
Dim LcTamanho, a As Long
'Verifica se Selecionou todos
If Len(Trim(GlCriterioSql)) > 0 Then
    LcCriSql = GlCriterioSql
Else
    If Len(Trim(Txt.Text)) = 0 Then
       LcCriSql = "select * From alid002 where RAZAOSOC like '*' order by RAZAOSOC"
      msg = "Aguarde, Criando Lista de Fornecedores ..."
    Else
      msg = "Aguarde, Filtrando Fornecedores Começados com " & UCase(Txt.Text)
      If Inicio Then
        LcCriSql = "select * From alid002 where RAZAOSOC like '" & UCase(Txt.Text) & "*' order by RAZAOSOC"
      Else
        LcCriSql = "select * From alid002 where RAZAOSOC like '*" & UCase(Txt.Text) & "*'  order by RAZAOSOC"
      End If
    End If
End If
'Set DbBase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
AbreBase
Set RsFornecedores = Dbbase.OpenRecordset(LcCriSql, dbOpenDynaset, dbSeeChanges, dbOptimistic) ', dbOpenDynaset)

LcTamanho = MostraCliente.Rows
a = 2
Me.Caption = msg
MostraCliente.Rows = 1

Do Until RsFornecedores.EOF
   If err.Number > 0 Then
      'MsgBox Err.Description
      Exit Do
   End If
   If Len(Trim(RsFornecedores!razaosoc)) > 0 Then
   If Not IsNull(RsFornecedores!razaosoc) Then
     
     MostraCliente.Rows = a
     MostraCliente.TextMatrix(a - 1, 0) = RsFornecedores!Codigo & ""
     MostraCliente.TextMatrix(a - 1, 1) = RsFornecedores!razaosoc & ""
     MostraCliente.TextMatrix(a - 1, 2) = RsFornecedores!CGC & ""
     MostraCliente.TextMatrix(a - 1, 3) = RTrim(RsFornecedores!End) & ""
     MostraCliente.TextMatrix(a - 1, 4) = RsFornecedores!fone1 & ""
     
     a = a + 1
     RsFornecedores.MoveNext
    End If
  End If
  
Loop
GlCriterioSql = ""
Me.Caption = "Fornecedores Começados com " & Txt.Text

RsFornecedores.Close
Set RsFornecedores = Nothing
MostraCliente.SetFocus


End Function

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
PesquisandoNota = False
GlFormA.SetFocus
End Sub

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
        FrmSaidaProduto.Txt(8).Text = MostraCliente.TextMatrix(a, 0)
        FrmSaidaProduto.Txt(9).Text = MostraCliente.TextMatrix(a, 1)
        Me.Visible = False
        FrmSaidaProduto.Txt(8).SetFocus
     Case Is = "FrmEntradaProduto"
        FrmEntradaProduto.Txt(8).Text = MostraCliente.TextMatrix(a, 0)
        FrmEntradaProduto.Txt(9).Text = MostraCliente.TextMatrix(a, 1)
        Me.Visible = False
        FrmEntradaProduto.CFOP.SetFocus
     Case Is = "Despesas"
        Despesas.Txt(2).SetFocus
        Despesas.Txt(2).Text = MostraCliente.TextMatrix(a, 0)
        Despesas.Txt(3).Text = MostraCliente.TextMatrix(a, 1)
        Me.Visible = False
     Case Is = "FrmPedido"
        FrmPedido.Txt(9).SetFocus
        FrmPedido.Txt(8).Text = MostraCliente.TextMatrix(a, 0)
        FrmPedido.Txt(9).Text = MostraCliente.TextMatrix(a, 1)
        Me.Visible = False
        
  End Select
  MostraCliente.BackColor = &H80000018
  PesquisandoNota = False
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
        FrmSaidaProduto.Txt(8).Text = MostraCliente.TextMatrix(a, 0)
        FrmSaidaProduto.Txt(9).Text = MostraCliente.TextMatrix(a, 1)
        Me.Visible = False
        FrmSaidaProduto.Txt(8).SetFocus
    Case Is = "Despesas"
        Despesas.Txt(2).SetFocus
        Despesas.Txt(2).Text = MostraCliente.TextMatrix(a, 0)
        Despesas.Txt(3).Text = MostraCliente.TextMatrix(a, 1)
        Me.Visible = False
    Case Is = "FrmEntradaProduto"
        FrmEntradaProduto.Txt(8).Text = MostraCliente.TextMatrix(a, 0)
        FrmEntradaProduto.Txt(9).Text = MostraCliente.TextMatrix(a, 1)
        Me.Visible = False
        FrmEntradaProduto.CFOP.SetFocus
    Case Is = "FrmPedido"
        FrmPedido.Txt(8).SetFocus
        FrmPedido.Txt(8).Text = MostraCliente.TextMatrix(a, 0)
        FrmPedido.Txt(9).Text = MostraCliente.TextMatrix(a, 1)
        Me.Visible = False
  End Select
  MostraCliente.BackColor = &H80000018
  PesquisandoNota = True
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

Private Sub Txt_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then
   MostraCliente.SetFocus
End If
If KeyCode = 27 Then Unload Me
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{C}"
End Sub


