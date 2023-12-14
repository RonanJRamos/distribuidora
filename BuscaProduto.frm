VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FrmBuscaProduto 
   BackColor       =   &H00CAE1A2&
   Caption         =   "Localiza Produto No Estoque"
   ClientHeight    =   6075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11310
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   ScaleHeight     =   6075
   ScaleWidth      =   11310
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5280
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Confirma F2"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5520
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar F10"
      Height          =   495
      Left            =   5640
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5520
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Mosta Todos Registros F3"
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5520
      Width           =   2415
   End
   Begin VB.TextBox txt 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   6855
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "BuscaProduto.frx":0000
      Height          =   4455
      Left            =   120
      OleObjectBlob   =   "BuscaProduto.frx":0014
      TabIndex        =   1
      Top             =   840
      Width           =   11055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Criterio"
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
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   885
   End
End
Attribute VB_Name = "FrmBuscaProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private a As Integer
Private Sub Command1_Click()
On Error Resume Next
Dim LcSql As String
LcSql = "select * from alid009 order by NOME"
Data1.RecordSource = LcSql
Data1.Refresh
Txt.SetFocus
End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 114 Then SendKeys "%+{M}"
If KeyCode = 113 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Command2_Click()
On Error Resume Next
Me.Visible = False

End Sub

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 114 Then SendKeys "%+{M}"
If KeyCode = 113 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Command3_Click()
On Error Resume Next
Dim bb As Database
Dim RsProduto As Recordset, RsUnidade As Recordset
LcCriSql = "select * from alid009"
LcSql1 = "Select * from alid004"

Set bb = OpenDatabase(GLBase, False, False) ' "dBASE III;")
Set RsProduto = bb.OpenRecordset(LcCriSql, dbOpenDynaset, dbSeeChanges, dbOptimistic) ', dbOpenDynaset)
Set RsUnidade = bb.OpenRecordset(LcSql1, dbOpenDynaset, dbSeeChanges, dbOptimistic) ', dbOpenDynaset)
 
 a = MostraCliente.Row
 Select Case GlFormA.Name
 
    Case Is = "FrmEntradaProduto"
        FrmEntradaProduto.Txt(2).Text = Data1.Recordset.Fields(1)
        FrmEntradaProduto.Txt(1).Text = Data1.Recordset.Fields(0)
        FrmEntradaProduto.Txt(4).Text = Data1.Recordset.Fields(7)
        ComNormal = RsProduto!QTDUNIMED
        LcCriterio = "cod='" & Data1.Recordset.Fields(0) & "'"
        RsProduto.FindFirst LcCriterio
        If Not RsProduto.NoMatch Then
           LccriterioUn = "cod='" & RsProduto!UNIMED & "'"
           RsUnidade.FindFirst LccriterioUn
           If Not RsUnidade.NoMatch Then
              FrmEntradaProduto.Unidade.Text = RsUnidade!Simbolo
           End If
        End If

   Case Is = "FrmSaidaProduto"
          FrmSaidaProduto.Txt(2).Text = Data1.Recordset.Fields(1)
        FrmSaidaProduto.Txt(1).Text = Data1.Recordset.Fields(0)
        FrmSaidaProduto.Txt(4).Text = Data1.Recordset.Fields(7)
        RsProduto.FindFirst "cod='" & Data1.Recordset.Fields(0) & "'"
        FrmSaidaProduto.valor(0).Text = Data1.Recordset.Fields(4)
        If Not IsNull(RsProduto!Ptab) Then PrecoVendaNormal = CCur(Data1.Recordset.Fields(4)) / RsProduto!QTDUNIMED Else PrecoVendaNormal = 0
        ComNormal = RsProduto!QTDUNIMED
        LcCriterio = "cod='" & Data1.Recordset.Fields(0) & "'"
        RsProduto.FindFirst LcCriterio
        If Not RsProduto.NoMatch Then
           LccriterioUn = "cod='" & RsProduto!UNIMED & "'"
           RsUnidade.FindFirst LccriterioUn
           If Not RsUnidade.NoMatch Then
              FrmSaidaProduto.Unidade.Text = RsUnidade!Simbolo
           End If
        End If
        FrmSaidaProduto.minimo.Text = RsProduto!MPVENDA & ""
        FrmSaidaProduto.cst.Text = RsProduto!cst
        If RsProduto!icms = 0 Then
            If Val(FrmSaidaProduto.cst.Text) = 60 Or Val(FrmSaidaProduto.cst.Text) = 160 Or Val(FrmSaidaProduto.cst.Text) = 260 Then
                    FrmSaidaProduto.icms.Text = "00"
            Else
              If Len(FrmSaidaProduto.Txt(5).Text) = 0 Then
                    FrmSaidaProduto.icms.Text = "18"
              Else
                    FrmSaidaProduto.icms = FrmSaidaProduto.Txt(5).Text
              End If
             End If
         Else
            FrmSaidaProduto.icms = RsProduto!icms
         End If
         If Not IsNull(RsProduto!MPVENDA) Then PrecoMimimodeVendaAlterado = CCur(Data1.Recordset.Fields(4)) / RsProduto!QTDUNIMED Else PrecoMimimodeVendaAlterado = 0
       ' FrmSaidaProduto.txt(5).Text = MostraCliente.TextMatrix(a, 5)
  
    Case Is = "FrmVales"
         FrmVales.Txt(2).Text = Data1.Recordset.Fields(1)
        FrmVales.Txt(1).Text = Data1.Recordset.Fields(0)
        FrmVales.Txt(4).Text = Data1.Recordset.Fields(7)
        RsProduto.FindFirst "cod='" & Data1.Recordset.Fields(0) & "'"
        FrmVales.valor(0).Text = Data1.Recordset.Fields(4)
        If Not IsNull(RsProduto!Ptab) Then PrecoVendaNormal = CCur(Data1.Recordset.Fields(4)) / RsProduto!QTDUNIMED Else PrecoVendaNormal = 0
        ComNormal = RsProduto!QTDUNIMED
        LcCriterio = "cod='" & Data1.Recordset.Fields(0) & "'"
        RsProduto.FindFirst LcCriterio
        If Not RsProduto.NoMatch Then
           LccriterioUn = "cod='" & RsProduto!UNIMED & "'"
           RsUnidade.FindFirst LccriterioUn
           If Not RsUnidade.NoMatch Then
              FrmVales.Unidade.Text = RsUnidade!Simbolo
           End If
        End If
        FrmVales.minimo.Text = RsProduto!MPVENDA & ""
        FrmVales.cst.Text = RsProduto!cst
        If RsProduto!icms = 0 Then
            If Val(FrmVales.cst.Text) = 60 Or Val(FrmVales.cst.Text) = 160 Or Val(FrmVales.cst.Text) = 260 Then
                    FrmVales.icms.Text = "00"
            Else
              If Len(FrmVales.Txt(5).Text) = 0 Then
                    FrmVales.icms.Text = "18"
              Else
                    FrmVales.icms = FrmVales.Txt(5).Text
              End If
             End If
         Else
            FrmVales.icms = RsProduto!icms
         End If
                 
         If Not IsNull(RsProduto!MPVENDA) Then PrecoMimimodeVendaAlterado = CCur(Data1.Recordset.Fields(4)) / RsProduto!QTDUNIMED Else PrecoMimimodeVendaAlterado = 0
       ' FrmSaidaProduto.txt(5).Text = MostraCliente.TextMatrix(a, 5)
    
    Case Is = "FrmReajustaPreco"
        FrmReajustaPreco.bo.Text = Data1.Recordset.Fields(1)
        FrmReajustaPreco.Codigo.Text = Data1.Recordset.Fields(0)
    Case Is = "ContratoFornecimento"
        ContratoFornecimento.Produto.Text = Data1.Recordset.Fields(1)
        ContratoFornecimento.CodProduto.Text = Data1.Recordset.Fields(0)
        ContratoFornecimento.ValorUnit.SetFocus
    Case Is = "FrmPedido"
        FrmPedido.Txt(2).Text = Data1.Recordset.Fields(1)
        FrmPedido.Txt(1).Text = Data1.Recordset.Fields(0)
        FrmPedido.Txt(4).Text = Data1.Recordset.Fields(7)
        RsProduto.FindFirst "cod='" & Data1.Recordset.Fields(0) & "'"
        'FrmPedido.Valor(0).Text = MostraCliente.TextMatrix(a, 2)
         ComNormal = RsProduto!QTDUNIMED
         FrmPedido.minimo.Text = RsProduto!MPVENDA & ""
         If Not IsNull(RsProduto!MPVENDA) Then PrecoMimimodeVendaAlterado = RsProduto!MPVENDA / RsProduto!QTDUNIMED Else PrecoMimimodeVendaAlterado = 0
       ' FrmSaidaProduto.txt(5).Text = MostraCliente.TextMatrix(a, 5)
        'FrmSaidaProduto.txt(5).Text = MostraC
    Case Is = "Orcamento"
        lcpesqunidade = "cod='" & Data1.Recordset.Fields(0) & "'"
        RsProduto.FindFirst lcpesqunidade
        If Not RsProduto.NoMatch Then
           orcamento.codigoproduto.Text = RsProduto!cod
           orcamento.NomeProduto.Text = RsProduto!Nome
           orcamento.ipi.Text = RsProduto!ipi & ""
           lcprocuraunidade = "cod='" & RsProduto!UNIMED & "'"
           orcamento.Industria.Text = RsProduto!fornecedor & ""
           Set RsUnidade = bb.OpenRecordset("select * From alid004 where COD='" & RsProduto!UNIMED & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)   ', dbOpenDynaset)
           If Not RsUnidade.EOF Then
              orcamento.Unidade.Text = RsUnidade!Simbolo
              orcamento.codigounidade.Text = RsUnidade!cod
           Else
              MsgBox "A unidade deste produto não Foi Cadastrada...", 64, "Aviso"
              orcamento.Unidade.Text = ""
           End If
           
           If Len(RsProduto!Ptab) > 0 Then
              orcamento.Unitario.Text = AcertaNumero(CStr(RsProduto!Ptab), GlDecimais)
              orcamento.preconormal.Text = RsProduto!Ptab
           Else
              orcamento.Unitario.Text = 0
           End If
        End If
        orcamento.Unidade.SetFocus
        RsProduto.Close
        RsUnidade.Close
     Case Is = "FrmProposta"
        FrmProposta.Txt(2).Text = Data1.Recordset.Fields(1)
        FrmProposta.Txt(1).Text = Data1.Recordset.Fields(0)
        FrmProposta.Txt(4).Text = Data1.Recordset.Fields(7)
        RsProduto.FindFirst "cod='" & Data1.Recordset.Fields(0) & "'"
        FrmProposta.valor(0).Text = Data1.Recordset.Fields(4)
        If Not IsNull(RsProduto!Ptab) Then PrecoVendaNormal = CCur(Data1.Recordset.Fields(4)) / RsProduto!QTDUNIMED Else PrecoVendaNormal = 0
        ComNormal = RsProduto!QTDUNIMED
        LcCriterio = "cod='" & Data1.Recordset.Fields(0) & "'"
        RsProduto.FindFirst LcCriterio
        If Not RsProduto.NoMatch Then
           LccriterioUn = "cod='" & RsProduto!UNIMED & "'"
           RsUnidade.FindFirst LccriterioUn
           If Not RsUnidade.NoMatch Then
              FrmProposta.Unidade.Text = RsUnidade!Simbolo
           End If
        End If
        FrmProposta.minimo.Text = RsProduto!MPVENDA & ""
         FrmProposta.cst.Text = RsProduto!cst
         
         If RsProduto!icms = 0 Then
            If Val(FrmProposta.cst.Text) = 60 Or Val(FrmProposta.cst.Text) = 160 Or Val(FrmProposta.cst.Text) = 260 Then
                    FrmProposta.icms.Text = "00"
            Else
              If Len(FrmProposta.Txt(5).Text) = 0 Then
                    FrmProposta.icms.Text = "18"
              Else
                    FrmProposta.icms = FrmVales.Txt(5).Text
              End If
             End If
         Else
            FrmProposta.icms = RsProduto!icms
         End If
         If Not IsNull(RsProduto!MPVENDA) Then PrecoMimimodeVendaAlterado = CCur(Data1.Recordset.Fields(4)) / RsProduto!QTDUNIMED Else PrecoMimimodeVendaAlterado = 0
       ' FrmSaidaProduto.txt(5).Text = MostraCliente.TextMatrix(a, 5)

    Case Is = "subitensproduto"
         subitensproduto.Codigo = Data1.Recordset.Fields("cod")
         subitensproduto.Produto = Data1.Recordset.Fields("nome")
         
    End Select
 Me.Visible = False
End Sub

Private Sub Command3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 114 Then SendKeys "%+{M}"
If KeyCode = 113 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub DBGrid1_DblClick()
SendKeys "%+{C}"
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 114 Then SendKeys "%+{M}"
If KeyCode = 113 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
Txt.Text = Txt.Text & Chr(KeyAscii)
End Sub

Private Sub DBGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
Dim LcLetras As String
Dim LcP As Integer
LcLetras = " ABCDEFGHIJLMNOPQRSTUVXZWY,./?\|[]{}:;~`0123456789+-*!@#$^&*()_="
'=== Vai Pesquisar para saber se é um Cracter valido
If KeyCode > 95 And KeyCode < 106 Then
   KeyCode = KeyCode - 48
End If
If KeyCode > 95 And KeyCode < 41 Then
   Exit Sub
End If
If KeyCode = 45 Then Exit Sub
If KeyCode = 106 Then Exit Sub
If KeyCode = 111 Then Exit Sub
If KeyCode = 109 Then Exit Sub
If KeyCode = 107 Then Exit Sub
If KeyCode = 116 Then Exit Sub
LcP = InStr(1, LcLetras, UCase(Chr(KeyCode)))

If KeyCode < 33 Or KeyCode > 40 Then
    If LcP > 0 Then
       Txt.Text = UCase(Txt.Text) & UCase(Chr(KeyCode))
    Else
       If KeyCode = 8 Then
          If Len(Txt.Text) > 0 Then
            Txt.Text = Left(Txt.Text, Len(Txt.Text) - 1)
          End If
       End If
    End If
End If
End Sub

Private Sub Form_Activate()
On Error Resume Next
Dim LcSql As String

If Len(GlCriterioSql) > 0 Then
   LcSql = GlCriterioSql
   GlCriterioSql = ""
   Txt.Text = Me.Tag
Else
   LcSql = "select * from alid009 order by NOME"
   Txt.Text = ""
End If

Data1.DatabaseName = GLBase
Data1.RecordSource = LcSql
Data1.Refresh
'txt.SetFocus
End Sub

Private Sub Form_Load()
On Error Resume Next

Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
GlFormA.SetFocus
End Sub

Private Sub Txt_Change()
On Error Resume Next
Dim LcCriterio As String
LCriterio = "nome like '" & Txt.Text & "*'"
Data1.Recordset.FindFirst LCriterio



End Sub

Private Sub Txt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 114 Then SendKeys "%+{M}"
If KeyCode = 113 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 40 Then
   DBGrid1.SetFocus
   SendKeys "{DOWN}"
End If
'MsgBox KeyCode
End Sub
