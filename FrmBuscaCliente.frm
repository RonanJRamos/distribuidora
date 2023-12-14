VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "dbgrid32.ocx"
Begin VB.Form FrmBuscaCliente 
   BackColor       =   &H00CAE1A2&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Localiza Cliente"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11640
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   11640
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   8160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5040
      Width           =   1140
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Mostra Todos Clientes  F3"
      Height          =   615
      Left            =   2160
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar F10"
      Height          =   615
      Left            =   4320
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Confirma F2"
      Default         =   -1  'True
      Height          =   615
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5040
      Width           =   1935
   End
   Begin VB.TextBox txt 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   6375
   End
   Begin MSDBGrid.DBGrid item 
      Bindings        =   "FrmBuscaCliente.frx":0000
      Height          =   3975
      Left            =   120
      OleObjectBlob   =   "FrmBuscaCliente.frx":0014
      TabIndex        =   1
      Top             =   720
      Width           =   11415
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Clientes Começados com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "FrmBuscaCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private a As Integer
Private Sub Command1_Click()
On Error Resume Next
Dim RsCl As Recordset
' a = MostraCliente.Row
 AbreBase
' a = MostraCliente.Row
Set RsVend = Dbbase.OpenRecordset("select * from alid200", dbOpenDynaset, dbSeeChanges, dbOptimistic)
  Select Case GlFormA.Name
     Case Is = "FrmSaidaProdutoAlternativo"
     If Not IsNull(Data2.Recordset.Fields(22)) Then LcCredito = Data2.Recordset.Fields(22) Else LcCredito = 0
            If Not IsNull(Data2.Recordset.Fields(23)) Then LcUtilizado = Data2.Recordset.Fields(23) Else LcUtilizado = 0
            If CCur(LcCredito) <= CCur(LcUtilizado) Then
                If FrmSaidaProdutoAlternativo.Natureza.Text <> "TRANSFERENCIA" Then
                    GlUtilizado = LcUtilizado
                    GlCredito = LcCredito
                    If Not GlNaoBloqueia Then
                       LiberacaoCli.Show
                       GlLibera = False
                       GlEscolha = True
                    
                       Do Until Not GlEscolha
                           DoEvents
                       Loop
                    Else
                       GlLibera = True
                    End If
                Else
                   GlLibera = True
                End If
                If Not GlLibera Then
                   FrmSaidaProdutoAlternativo.txt(9).Text = ""
                   FrmSaidaProdutoAlternativo.txt(9).SetFocus
                Else
                   FrmSaidaProdutoAlternativo.limite.Text = LcCredito
                   FrmSaidaProdutoAlternativo.utilizado.Text = LcUtilizado
                   FrmSaidaProdutoAlternativo.txt(8).Text = Data2.Recordset.Fields(0)
                   FrmSaidaProdutoAlternativo.txt(9).Text = Data2.Recordset.Fields(1)
                End If
                
            Else
                FrmSaidaProdutoAlternativo.limite.Text = LcCredito
                FrmSaidaProdutoAlternativo.utilizado.Text = LcUtilizado
                FrmSaidaProdutoAlternativo.txt(8).Text = Data2.Recordset.Fields(0)
                FrmSaidaProdutoAlternativo.txt(9).Text = Data2.Recordset.Fields(1)
                If UCase(Data2.Recordset.Fields(4)) = "MG" Then
                    ClienteForaEstado = False
                Else
                    ClienteForaEstado = True
                End If
                FrmSaidaProdutoAlternativo.txt(8).SetFocus
            End If
            '===> Busca o Vendedor
            LcBusV = "Nome='" & Data2.Recordset.Fields("TelemarketingAtende") & "'"
            RsVend.FindFirst LcBusV
            If Not RsVend.NoMatch Then
               FrmSaidaProdutoAlternativo.txt(7).Text = RsVend!Nome & ""
               FrmSaidaProdutoAlternativo.txt(10).Text = RsVend!Codigo & ""
            Else
               FrmSaidaProdutoAlternativo.txt(7).Text = Data2.Recordset.Fields("TelemarketingAtende") & ""
               FrmSaidaProdutoAlternativo.txt(10).Text = ""
            End If
            If VerificaAtraso(Data2.Recordset.Fields("codigo")) Then
               FrmSaidaProdutoAlternativo.txt(9).SetFocus
            End If
            
            Me.Visible = False
            'FrmSaidaProduto.verificavale
     Case Is = "FrmSaidaProduto"
            If Not IsNull(Data2.Recordset.Fields(22)) Then LcCredito = Data2.Recordset.Fields(22) Else LcCredito = 0
            If Not IsNull(Data2.Recordset.Fields(23)) Then LcUtilizado = Data2.Recordset.Fields(23) Else LcUtilizado = 0
            If CCur(LcCredito) <= CCur(LcUtilizado) Then
                If FrmSaidaProduto.Natureza.Text <> "TRANSFERENCIA" Then
                    GlUtilizado = LcUtilizado
                    GlCredito = LcCredito
                    If Not GlNaoBloqueia Then
                       LiberacaoCli.Show
                       GlLibera = False
                       GlEscolha = True
                    
                       Do Until Not GlEscolha
                           DoEvents
                       Loop
                    Else
                       GlLibera = True
                    End If
                Else
                   GlLibera = True
                End If
                If Not GlLibera Then
                   FrmSaidaProduto.txt(9).Text = ""
                   FrmSaidaProduto.txt(9).SetFocus
                Else
                   FrmSaidaProduto.limite.Text = LcCredito
                   FrmSaidaProduto.utilizado.Text = LcUtilizado
                   FrmSaidaProduto.txt(8).Text = Data2.Recordset.Fields(0)
                   FrmSaidaProduto.txt(9).Text = Data2.Recordset.Fields(1)
                End If
                
            Else
                FrmSaidaProduto.limite.Text = LcCredito
                FrmSaidaProduto.utilizado.Text = LcUtilizado
                FrmSaidaProduto.txt(8).Text = Data2.Recordset.Fields(0)
                FrmSaidaProduto.txt(9).Text = Data2.Recordset.Fields(1)
                If UCase(Data2.Recordset.Fields(4)) = "MG" Then
                    ClienteForaEstado = False
                Else
                    ClienteForaEstado = True
                End If
                FrmSaidaProduto.txt(8).SetFocus
            End If
            '===> Busca o Vendedor
            LcBusV = "Nome='" & Data2.Recordset.Fields("TelemarketingAtende") & "'"
            RsVend.FindFirst LcBusV
            If Not RsVend.NoMatch Then
               FrmSaidaProduto.txt(7).Text = RsVend!Nome & ""
               FrmSaidaProduto.txt(10).Text = RsVend!Codigo & ""
            Else
               FrmSaidaProduto.txt(7).Text = Data2.Recordset.Fields("TelemarketingAtende") & ""
               FrmSaidaProduto.txt(10).Text = ""
            End If
            If VerificaAtraso(Data2.Recordset.Fields("codigo")) Then
               FrmSaidaProduto.txt(9).SetFocus
            End If
            
            Me.Visible = False
            FrmSaidaProduto.verificavale

   Case Is = "FrmVales"
            If Not IsNull(Data2.Recordset.Fields(22)) Then LcCredito = Data2.Recordset.Fields(22) Else LcCredito = 0
            If Not IsNull(Data2.Recordset.Fields(23)) Then LcUtilizado = Data2.Recordset.Fields(23) Else LcUtilizado = 0
            Set RsCl = Dbbase.OpenRecordset("Select * from alid001 where codigo= '" & Data2.Recordset.Fields("codigo") & "'")
            If Not RsCl.EOF Then
               If RsCli!bloqueado Then
                   MsgBox "Cliente Bloqueado!", 64, "Aviso"
                   Exit Sub
               End If
            End If
            
            If CCur(LcCredito) <= CCur(LcUtilizado) Then
                If FrmVales.Natureza.Text <> "TRANSFERENCIA" Then
                    GlUtilizado = LcUtilizado
                    GlCredito = LcCredito
                    If Not GlNaoBloqueia Then
                       LiberacaoCli.Show
                       GlLibera = False
                       GlEscolha = True
                    
                       Do Until Not GlEscolha
                           DoEvents
                       Loop
                    Else
                       GlLibera = True
                    End If
                Else
                   GlLibera = True
                End If
                If Not GlLibera Then
                   FrmVales.txt(9).Text = ""
                   FrmVales.txt(9).SetFocus
                Else
                   FrmVales.limite.Text = LcCredito
                   FrmVales.utilizado.Text = LcUtilizado
                   FrmVales.txt(8).Text = Data2.Recordset.Fields(0)
                   FrmVales.txt(9).Text = Data2.Recordset.Fields(1)
                End If
                
            Else
                FrmVales.limite.Text = LcCredito
                FrmVales.utilizado.Text = LcUtilizado
                FrmVales.txt(8).Text = Data2.Recordset.Fields(0)
                FrmVales.txt(9).Text = Data2.Recordset.Fields(1)
                
                FrmVales.txt(8).SetFocus
            End If
            '===> Busca o Vendedor
            LcBusV = "Nome='" & Data2.Recordset.Fields("TelemarketingAtende") & "'"
            RsVend.FindFirst LcBusV
            If Not RsVend.NoMatch Then
               FrmVales.txt(7).Text = RsVend!Nome & ""
               FrmVales.txt(10).Text = RsVend!Codigo & ""
            Else
               FrmVales.txt(7).Text = Data2.Recordset.Fields("TelemarketingAtende") & ""
               FrmVales.txt(10).Text = ""
            End If
            If VerificaAtraso(Data2.Recordset.Fields("codigo")) Then
               FrmVales.txt(9).SetFocus
            End If
            
            Me.Visible = False

    Case Is = "Receitas"
        Receitas.txt(2).SetFocus
        Receitas.txt(2).Text = Data2.Recordset.Fields(0)
        Receitas.txt(3).Text = Data2.Recordset.Fields(1)
        Me.Visible = False
   Case Is = "Orcamento"
        Orcamento.codigoproduto.SetFocus
        Orcamento.CodigoCliente.Text = Data2.Recordset.Fields(0)
        Orcamento.NomeCliente.Text = Data2.Recordset.Fields(1)
        Orcamento.LimiteCredito.Text = Data2.Recordset.Fields(22)
        Orcamento.LimiteUtilizado.Text = Data2.Recordset.Fields(23)
        Me.Visible = False
    Case Is = "FrmPedido"
        FrmPedido.txt(18).SetFocus
        FrmPedido.txt(18).Text = Data2.Recordset.Fields(0)
        FrmPedido.txt(17).Text = Data2.Recordset.Fields(1)
        Me.Visible = False
     Case Is = "ContratoFornecimento"
        ContratoFornecimento.Contrato.SetFocus
        ContratoFornecimento.CodCliente.Text = Data2.Recordset.Fields(0)
        ContratoFornecimento.Cliente.Text = Data2.Recordset.Fields(1)
        Me.Visible = False
     Case Is = "Frmcheques"
        Frmcheques.txt(1).SetFocus
        Frmcheques.Codigo.Text = Data2.Recordset.Fields(0)
        Frmcheques.txt(1).Text = Data2.Recordset.Fields(1)
        Me.Visible = False
     Case Is = "FrmProposta"
          If Not IsNull(Data2.Recordset.Fields(22)) Then LcCredito = Data2.Recordset.Fields(22) Else LcCredito = 0
           
            '===> Busca o Vendedor
            LcBusV = "Nome='" & Data2.Recordset.Fields("TelemarketingAtende") & "'"
            Set RsCl = Dbbase.OpenRecordset("Select * from alid001 where codigo= '" & Data2.Recordset.Fields("codigo") & "'")
            If Not RsCl.EOF Then
               If RsCl!comodato Then
                  FrmProposta.comodato.Value = 1
               Else
                  FrmProposta.comodato.Value = 0
               End If
               If RsCl!bloqueado Then
                   MsgBox "Cliente bloqueado!", 64, "Aviso"
                   Exit Sub
               End If
            End If
            
            RsCl.Close
            Set RsCl = Nothing
            RsVend.FindFirst LcBusV
            If Not RsVend.NoMatch Then
               FrmProposta.txt(8).Text = Data2.Recordset.Fields("codigo")
               FrmProposta.txt(9).Text = Data2.Recordset.Fields("razaosoc")
               FrmProposta.txt(7).Text = RsVend!Nome & ""
               FrmProposta.txt(10).Text = RsVend!Codigo & ""
            Else
               FrmProposta.txt(8).Text = Data2.Recordset.Fields("codigo")
               FrmProposta.txt(9).Text = Data2.Recordset.Fields("razaosoc")
               FrmProposta.txt(7).Text = Data2.Recordset.Fields("TelemarketingAtende") & ""
               FrmProposta.txt(10).Text = ""
            End If
            
            Me.Visible = False
            'FrmProposta.txt(1).SetFocus
  End Select
RsVend.Close
Dbbase.Close
Set RsVend = Nothing
Set bbase = Nothing

  LcCarr = False
  MostraCliente.BackColor = &H80000018
  GlFormA.SetFocus


End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{C}"
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 114 Then SendKeys "%+{M}"
If KeyCode = 121 Then SendKeys "%+{F}"

End Sub

Private Sub Command2_Click()
On Error Resume Next
Me.Visible = False
  Select Case GlFormA.Name
     Case Is = "FrmSaidaProduto"
        FrmSaidaProduto.txt(9).SetFocus
     Case Is = "Receitas"
        Receitas.txt(3).SetFocus
     Case Is = "Orcamento"
        Orcamento.NomeCliente.SetFocus
     Case Is = "FrmPedido"
       
     Case Is = "Frmcheques"
        Frmcheques.txt(1).SetFocus
     Case Is = "FrmProposta"
        FrmProposta.txt(9).SetFocus
     Case Is = "ContratoFornecimento"
        ContratoFornecimento.Cliente.SetFocus
  End Select

End Sub

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{C}"
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 114 Then SendKeys "%+{M}"
If KeyCode = 121 Then SendKeys "%+{F}"

End Sub

Private Sub Command3_Click()
On Error Resume Next
LcSql = "SELECT ALID001.codigo,ALID001.RAZAOSOC, ALID001.END, ALID001.BAIRRO, ALID001.ESTADO, ALID001.CEP, ALID001.FONE1, ALID001.FONE2, ALID001.FAX, ALID001.CONTATO, ALID001.CGC, ALID001.INSCEST, ALID001.ULTCOMPRA, ALID001.MEDIA, ALID001.CATEGC, ALID001.CATEGF, ALID001.ENDCOB, ALID001.BAIRROCOB, ALID001.CIDADECOB, ALID001.ESTADOCOB, ALID001.CEPCOB, ALID001.TelemarketingAtende, ALID001.LimiteCredito, ALID001.CreditoUtilizado, ALID001.CondicaoEspecial, ALID005.NOME "
LcSql = LcSql & "FROM ALID001 INNER JOIN ALID005 ON ALID001.CIDADE = ALID005.COD order by razaosoc"
Data2.DatabaseName = GLBase
Data2.RecordSource = LcSql
Data2.Refresh
txt.SetFocus
End Sub

Private Sub Command3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{C}"
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 114 Then SendKeys "%+{M}"
If KeyCode = 121 Then SendKeys "%+{F}"

End Sub

Private Sub Form_Activate()
On Error Resume Next
Dim LcSql  As String
item.Columns(3).DataField = "end"
item.Columns(4).DataField = "nome"
If Len(GlCriterioSql) = 0 Then
   LcSql = "SELECT ALID001.codigo,ALID001.RAZAOSOC, ALID001.END, ALID001.BAIRRO, ALID001.ESTADO, ALID001.CEP, ALID001.FONE1, ALID001.FONE2, ALID001.FAX, ALID001.CONTATO, ALID001.CGC, ALID001.INSCEST, ALID001.ULTCOMPRA, ALID001.MEDIA, ALID001.CATEGC, ALID001.CATEGF, ALID001.ENDCOB, ALID001.BAIRROCOB, ALID001.CIDADECOB, ALID001.ESTADOCOB, ALID001.CEPCOB, ALID001.TelemarketingAtende, ALID001.LimiteCredito, ALID001.CreditoUtilizado, ALID001.CondicaoEspecial, alid001.cpf, ALID005.NOME "
   LcSql = LcSql & "FROM ALID001 INNER JOIN ALID005 ON ALID001.CIDADE = ALID005.COD order by razaosoc"
Else
   LcSql = "SELECT ALID001.codigo,ALID001.RAZAOSOC, ALID001.END, ALID001.BAIRRO, ALID001.ESTADO, ALID001.CEP, ALID001.FONE1, ALID001.FONE2, ALID001.FAX, ALID001.CONTATO, ALID001.CGC, ALID001.INSCEST, ALID001.ULTCOMPRA, ALID001.MEDIA, ALID001.CATEGC, ALID001.CATEGF, ALID001.ENDCOB, ALID001.BAIRROCOB, ALID001.CIDADECOB, ALID001.ESTADOCOB, ALID001.CEPCOB, ALID001.TelemarketingAtende, ALID001.LimiteCredito, ALID001.CreditoUtilizado, ALID001.CondicaoEspecial, alid001.cpf, ALID005.NOME "
   LcSql = LcSql & "FROM ALID001 INNER JOIN ALID005 ON ALID001.CIDADE = ALID005.COD "
   LcSql = LcSql & GlCriterioSql
End If
'MsgBox LcSql
Data2.DatabaseName = GLBase
Data2.RecordSource = LcSql
Data2.Refresh

End Sub



Private Sub Item_DblClick()
SendKeys "%+{C}"
End Sub

Private Sub item_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{C}"
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 114 Then SendKeys "%+{M}"
If KeyCode = 121 Then SendKeys "%+{F}"

End Sub

Private Sub Txt_Change()
On Error Resume Next
Dim LcCriterio As String
LcCriterio = "razaosoc like '" & txt.Text & "*'"
Data2.Recordset.FindFirst LcCriterio
End Sub

Private Sub Txt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{C}"
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 114 Then SendKeys "%+{M}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 40 Then
   item.SetFocus
   SendKeys "{DOWN}"
End If
End Sub
