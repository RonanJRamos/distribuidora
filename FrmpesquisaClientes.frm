VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmPesquisaCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pesquisa Clientes"
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
      Height          =   3975
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   7011
      _Version        =   393216
      Rows            =   17
      Cols            =   8
      FixedCols       =   0
      BackColor       =   -2147483624
      FocusRect       =   2
      SelectionMode   =   1
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
Attribute VB_Name = "FrmPesquisaCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private LcCorAnterior, LcCarr, a As Integer
Private Sub CmdCancelar_Click()
On Error Resume Next
Me.Visible = False
LcCarr = False
FrmLocacao.Txt(4).SetFocus
 
End Sub

Private Sub CmdCancelar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub CmdOk_Click()
ExibePesquisa
End Sub

Private Sub CmdOk_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Form_Activate()
Me.Refresh
If Len(GlCriterioSql) = 0 Then
   'Txt.Text = ""
   GlCriterioSql = ""
   
End If
If Not LcCarr Then
   ExibePesquisa
   LcCarr = True
End If
End Sub

Private Sub Form_Load()
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
GeraGrid
ExibePesquisa
End Sub
Function GeraGrid()
MostraCliente.ColAlignment(0) = 7
MostraCliente.ColAlignment(1) = 1
MostraCliente.ColAlignment(2) = 1
MostraCliente.ColAlignment(3) = 1
MostraCliente.ColAlignment(4) = 1
MostraCliente.ColAlignment(7) = 1

MostraCliente.ColWidth(0) = 0
MostraCliente.ColWidth(1) = 3600
MostraCliente.ColWidth(2) = 1500
MostraCliente.ColWidth(3) = 3800
MostraCliente.ColWidth(4) = 1500
MostraCliente.ColWidth(5) = 0
MostraCliente.ColWidth(6) = 0
MostraCliente.ColWidth(7) = 1300
MostraCliente.TextMatrix(0, 0) = "Código"
MostraCliente.TextMatrix(0, 1) = "Nome"
MostraCliente.TextMatrix(0, 2) = "C.N.P.J/C.P.F."
MostraCliente.TextMatrix(0, 3) = "Endereço"
MostraCliente.TextMatrix(0, 7) = "Fone"
MostraCliente.TextMatrix(0, 4) = "Cidade"
LcTamanhoGrid = 1
End Function
Function ExibePesquisa()
On Error GoTo errorExibeCli
Dim rsCliente As Recordset, RsCliente1 As Recordset, RsCliente2 As Recordset, RsCidade As Recordset
Dim LcCriSql, LcCriSql1, LcCriSql2 As String
Dim LcTamanho, a As Long
Dim LcAchou As Integer

'Verifica se Selecionou todos
If Len(Trim(GlCriterioSql)) > 0 Then
    LcCriSql = GlCriterioSql
    
Else
    If Len(Trim(Txt.Text)) = 0 Then
       LcCriSql = "select * From alid001 where RAZAOSOC like '*' order by RAZAOSOC"
      Msg = "Aguarde, Criando Lista de Clientes..."
    Else
      Msg = "Aguarde, Filtrando Clientes Começados com " & UCase(Txt.Text)
      If Inicio Then
        LcCriSql = "select * From alid001 where RAZAOSOC like '" & UCase(Txt.Text) & "*' order by RAZAOSOC"
      Else
        LcCriSql = "select * From alid001 where RAZAOSOC like '*" & UCase(Txt.Text) & "*'  order by RAZAOSOC"
      End If
    End If
End If
'Set DbBase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
AbreBase
Set rsCliente = Dbbase.OpenRecordset(LcCriSql, dbOpenDynaset, dbSeeChanges, dbOptimistic) ', dbOpenDynaset)
Set RsCidade = Dbbase.OpenRecordset("alid005", dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcTamanho = MostraCliente.Rows
a = 2
Me.Caption = Msg
MostraCliente.Rows = 1
LcAchou = False
Do Until rsCliente.EOF
  LcAchou = True
  If Len(Trim(rsCliente!RAZAOSOC)) > 0 Then
   If Not IsNull(rsCliente!RAZAOSOC) Then
     LcCidade = "cod='" & rsCliente!cidade & "'"
     RsCidade.FindFirst LcCidade
     MostraCliente.Rows = a
     MostraCliente.TextMatrix(a - 1, 0) = rsCliente!Codigo & ""
     MostraCliente.TextMatrix(a - 1, 1) = rsCliente!RAZAOSOC & ""
     If rsCliente!CGC <> "  .   .   /    -  " Then
        MostraCliente.TextMatrix(a - 1, 2) = rsCliente!CGC & ""
     Else
        MostraCliente.TextMatrix(a - 1, 2) = rsCliente!cpf & ""
     End If
     MostraCliente.TextMatrix(a - 1, 3) = RTrim(rsCliente!End) & ""
     MostraCliente.TextMatrix(a - 1, 7) = rsCliente!Fone1 & ""
     MostraCliente.TextMatrix(a - 1, 5) = rsCliente!LimiteCredito & ""
     MostraCliente.TextMatrix(a - 1, 6) = rsCliente!CreditoUtilizado & ""
     LcCidade = "cod='" & rsCliente!cidade & "'"
     RsCidade.FindFirst LcCidade
     If Not RsCidade.EOF Then
        MostraCliente.TextMatrix(a - 1, 4) = RsCidade!Nome & ""
     End If
     If rsCliente!bloqueado Then
          MostraCliente.Row = a - 1
          
          'Seleciona até a última coluna
          For x = 1 To MostraCliente.Cols - 1
            'Aplica a cor
            MostraCliente.Col = x
            MostraCliente.CellBackColor = vbRed
          Next
          
     End If
     a = a + 1
     rsCliente.MoveNext
    End If
  End If
  
Loop
If Not LcAchou Then GlCriterioSql = ""

Me.Caption = "Clientes Começados com " & Txt.Text
GlCriterioSql = ""
rsCliente.Close
Set rsCliente = Nothing
MostraCliente.SetFocus
Exit Function
errorExibeCli:
If err = 5 Then Resume Next
If err = 3420 Then
   AbreBanco (LcTabl)
Else
   If err = 3021 Then
      Resume Next
   Else
   '   MsgBox Err.Description & " " & Err
   End If
  ' Resume 0
End If

End Function

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
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
Dim a As Long
 On Error Resume Next
 a = MostraCliente.Row
  Select Case GlFormA.Name
  Case Is = "alid015"
        alid015.BotoesClienteAnterior.Text = alid015.Cliente.Text
         alid015.Cliente.Text = MostraCliente.TextMatrix(a, 0)
         alid015.Nome.Text = MostraCliente.TextMatrix(a, 1)
         alid015.botoes.Buttons(1).Enabled = True
         Unload Me
   Case Is = "FrmDuplicaPedido"
            FrmDuplicaPedido.CodCliente.Text = MostraCliente.TextMatrix(a, 0)
            FrmDuplicaPedido.NomeCliente.Text = MostraCliente.TextMatrix(a, 1)
            Unload Me
     Case Is = "FrmSaidaProduto"
            If Not IsNull(MostraCliente.TextMatrix(a, 5)) Then LcCredito = MostraCliente.TextMatrix(a, 5) Else LcCredito = 0
            If Not IsNull(MostraCliente.TextMatrix(a, 5)) Then LcUtilizado = MostraCliente.TextMatrix(a, 6) Else LcUtilizado = 0
            If CCur(LcCredito) <= CCur(LcUtilizado) Then
                GlUtilizado = LcUtilizado
                GlCredito = LcCredito
                LiberacaoCli.Show
                GlLibera = False
                GlEscolha = True
                Do Until Not GlEscolha
                    DoEvents
                Loop
                If Not GlLibera Then
                   FrmSaidaProduto.Txt(9).Text = ""
                   FrmSaidaProduto.Txt(9).SetFocus
                Else
                   FrmSaidaProduto.limite.Text = LcCredito
                   FrmSaidaProduto.utilizado.Text = LcUtilizado
                   FrmSaidaProduto.Txt(8).Text = MostraCliente.TextMatrix(a, 0)
                   FrmSaidaProduto.Txt(9).Text = MostraCliente.TextMatrix(a, 1)
                End If
            Else
                limite.Text = LcCredito
                utilizado.Text = LcUtilizado
                FrmSaidaProduto.Txt(8).Text = MostraCliente.TextMatrix(a, 0)
                FrmSaidaProduto.Txt(9).Text = MostraCliente.TextMatrix(a, 1)
                FrmSaidaProduto.Txt(8).SetFocus
            End If
            Me.Visible = False
            If VerificaAtraso(FrmSaidaProduto.Txt(8).Text) Then
               FrmSaidaProduto.Txt(9).SetFocus
            End If
            
    Case Is = "Receitas"
        Receitas.Txt(2).SetFocus
        Receitas.Txt(2).Text = MostraCliente.TextMatrix(a, 0)
        Receitas.Txt(3).Text = MostraCliente.TextMatrix(a, 1)
        Me.Visible = False
   Case Is = "Orcamento"
        orcamento.codigoproduto.SetFocus
        orcamento.CodigoCliente.Text = MostraCliente.TextMatrix(a, 0)
        orcamento.NomeCliente.Text = MostraCliente.TextMatrix(a, 1)
        orcamento.LimiteCredito.Text = MostraCliente.TextMatrix(a, 5)
        orcamento.LimiteUtilizado.Text = MostraCliente.TextMatrix(a, 6)
        Me.Visible = False
    Case Is = "FrmPedido"
        FrmPedido.Txt(18).SetFocus
        FrmPedido.Txt(18).Text = MostraCliente.TextMatrix(a, 0)
        FrmPedido.Txt(17).Text = MostraCliente.TextMatrix(a, 1)
        Me.Visible = False
     Case Is = "Frmcheques"
        Frmcheques.Txt(1).SetFocus
        Frmcheques.Codigo.Text = MostraCliente.TextMatrix(a, 0)
        Frmcheques.Txt(1).Text = MostraCliente.TextMatrix(a, 1)
        Me.Visible = False
     Case Is = "FrmProposta"
         If Not IsNull(MostraCliente.TextMatrix(a, 5)) Then LcCredito = MostraCliente.TextMatrix(a, 5) Else LcCredito = 0
            If Not IsNull(MostraCliente.TextMatrix(a, 5)) Then LcUtilizado = MostraCliente.TextMatrix(a, 6) Else LcUtilizado = 0
            If bloqueado(MostraCliente.TextMatrix(a, 0)) Then
                MsgBox "O cliente esta bloqueado!", 64, "Aviso"
                Exit Sub
            End If
            If CCur(LcCredito) <= CCur(LcUtilizado) Then
                GlUtilizado = LcUtilizado
                GlCredito = LcCredito
                LiberacaoCli.Show
                GlLibera = False
                GlEscolha = True
                Do Until Not GlEscolha
                    DoEvents
                Loop
                If Not GlLibera Then
                   FrmProposta.Txt(9).Text = ""
                   FrmProposta.Txt(9).SetFocus
                Else
                   FrmProposta.limite.Text = LcCredito
                   FrmProposta.utilizado.Text = LcUtilizado
                   FrmProposta.Txt(8).Text = MostraCliente.TextMatrix(a, 0)
                   FrmProposta.Txt(9).Text = MostraCliente.TextMatrix(a, 1)
                End If
            Else
                limite.Text = LcCredito
                utilizado.Text = LcUtilizado
                FrmProposta.Txt(8).Text = MostraCliente.TextMatrix(a, 0)
                FrmProposta.Txt(9).Text = MostraCliente.TextMatrix(a, 1)
                FrmProposta.Txt(8).SetFocus
            End If
            Me.Visible = False
            If VerificaAtraso(FrmSaidaProduto.Txt(8).Text) Then
               FrmProposta.Txt(9).SetFocus
            End If

  End Select
  LcCarr = False
  MostraCliente.BackColor = &H80000018
  GlFormA.SetFocus
End Sub
Function bloqueado(Codigo As Long) As Boolean
 Dim rsCliente As Recordset
 AbreBase
 LcCriSql = "Select * from alid001 where codigo='" & Right("00000" & Codigo, 5) & "'"
 Set rsCliente = Dbbase.OpenRecordset(LcCriSql, dbOpenDynaset, dbSeeChanges, dbOptimistic)
If Not rsCliente.EOF Then
  bloqueado = rsCliente!bloqueado
End If
rsCliente.Close
Set rsCliente = Nothing

End Function
Private Sub MostraCliente_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Dim a As Long
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then
   
   a = MostraCliente.Row
  Select Case GlFormA.Name
    Case Is = "alid015"
        alid015.BotoesClienteAnterior.Text = alid015.Cliente.Text
         alid015.Cliente.Text = MostraCliente.TextMatrix(a, 0)
         alid015.Nome.Text = MostraCliente.TextMatrix(a, 1)
         alid015.botoes.Buttons(1).Enabled = True
         Unload Me
    Case Is = "FrmDuplicaPedido"
            FrmDuplicaPedido.CodCliente.Text = MostraCliente.TextMatrix(a, 0)
            FrmDuplicaPedido.NomeCliente.Text = MostraCliente.TextMatrix(a, 1)
            Unload Me
     Case Is = "FrmSaidaProduto"
         If Not IsNull(MostraCliente.TextMatrix(a, 5)) Then LcCredito = MostraCliente.TextMatrix(a, 5) Else LcCredito = 0
            If Not IsNull(MostraCliente.TextMatrix(a, 5)) Then LcUtilizado = MostraCliente.TextMatrix(a, 6) Else LcUtilizado = 0
            If CCur(LcCredito) <= CCur(LcUtilizado) Then
                GlUtilizado = LcUtilizado
                GlCredito = LcCredito
                LiberacaoCli.Show
                GlLibera = False
                GlEscolha = True
                Do Until Not GlEscolha
                    DoEvents
                Loop
                If Not GlLibera Then
                   FrmSaidaProduto.Txt(9).Text = ""
                   FrmSaidaProduto.Txt(9).SetFocus
                   
                Else
                   FrmSaidaProduto.limite.Text = LcCredito
                   FrmSaidaProduto.utilizado.Text = LcUtilizado
                   FrmSaidaProduto.Txt(8).Text = MostraCliente.TextMatrix(a, 0)
                   FrmSaidaProduto.Txt(9).Text = MostraCliente.TextMatrix(a, 1)
                End If
            Else
                limite.Text = LcCredito
                utilizado.Text = LcUtilizado
                FrmSaidaProduto.Txt(8).Text = MostraCliente.TextMatrix(a, 0)
                FrmSaidaProduto.Txt(9).Text = MostraCliente.TextMatrix(a, 1)
                FrmSaidaProduto.Txt(8).SetFocus
            End If
            Me.Visible = False
    Case Is = "Receitas"
        Receitas.Txt(2).SetFocus
        Receitas.Txt(2).Text = MostraCliente.TextMatrix(a, 0)
        Receitas.Txt(3).Text = MostraCliente.TextMatrix(a, 1)
        Me.Visible = False
    Case Is = "Orcamento"
        orcamento.CodigoCliente.Text = MostraCliente.TextMatrix(a, 0)
        orcamento.NomeCliente.Text = MostraCliente.TextMatrix(a, 1)
        orcamento.LimiteCredito.Text = MostraCliente.TextMatrix(a, 5)
        orcamento.LimiteUtilizado.Text = MostraCliente.TextMatrix(a, 6)
        orcamento.codigoproduto.SetFocus
        Me.Visible = False
    Case Is = "FrmPedido"
        FrmPedido.Txt(18).SetFocus
        FrmPedido.Txt(18).Text = MostraCliente.TextMatrix(a, 0)
        FrmPedido.Txt(17).Text = MostraCliente.TextMatrix(a, 1)
        Me.Visible = False
    Case Is = "Frmcheques"
        Frmcheques.Txt(1).SetFocus
        Frmcheques.Codigo.Text = MostraCliente.TextMatrix(a, 0)
        Frmcheques.Txt(1).Text = MostraCliente.TextMatrix(a, 1)
        Me.Visible = False
    Case Is = "FrmProposta"
         If Not IsNull(MostraCliente.TextMatrix(a, 5)) Then LcCredito = MostraCliente.TextMatrix(a, 5) Else LcCredito = 0
            If Not IsNull(MostraCliente.TextMatrix(a, 5)) Then LcUtilizado = MostraCliente.TextMatrix(a, 6) Else LcUtilizado = 0
            If bloqueado(MostraCliente.TextMatrix(a, 0)) Then
                MsgBox "O cliente esta bloqueado!", 64, "Aviso"
                Exit Sub
            End If
            If CCur(LcCredito) <= CCur(LcUtilizado) Then
                GlUtilizado = LcUtilizado
                GlCredito = LcCredito
                LiberacaoCli.Show
                GlLibera = False
                GlEscolha = True
                Do Until Not GlEscolha
                    DoEvents
                Loop
                If Not GlLibera Then
                   FrmProposta.Txt(9).Text = ""
                   FrmProposta.Txt(9).SetFocus
                Else
                   FrmProposta.limite.Text = LcCredito
                   FrmProposta.utilizado.Text = LcUtilizado
                   FrmProposta.Txt(8).Text = MostraCliente.TextMatrix(a, 0)
                   FrmProposta.Txt(9).Text = MostraCliente.TextMatrix(a, 1)
                End If
            Else
                limite.Text = LcCredito
                utilizado.Text = LcUtilizado
                FrmProposta.Txt(8).Text = MostraCliente.TextMatrix(a, 0)
                FrmProposta.Txt(9).Text = MostraCliente.TextMatrix(a, 1)
                FrmProposta.Txt(8).SetFocus
            End If
            
            Me.Visible = False
            If VerificaAtraso(FrmSaidaProduto.Txt(8).Text) Then
               FrmProposta.Txt(9).SetFocus
            End If


  End Select
  MostraCliente.BackColor = &H80000018
  LcCarr = False
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


