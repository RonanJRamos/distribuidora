VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form GastosFixos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gastos Fixos"
   ClientHeight    =   5340
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSDBGrid.DBGrid item 
      Bindings        =   "GastosFixos.frx":0000
      Height          =   4215
      Left            =   120
      OleObjectBlob   =   "GastosFixos.frx":0014
      TabIndex        =   0
      Top             =   360
      Width           =   5415
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK F2"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Fechar F10"
      Height          =   615
      Left            =   1800
      TabIndex        =   2
      Top             =   4680
      Width           =   1575
   End
End
Attribute VB_Name = "GastosFixos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private a As Integer
Option Explicit

Private Sub Command1_Click()
Data1.Refresh
item.Refresh
End Sub

Private Sub CancelButton_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub CancelButton_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Form_Activate()
On Error Resume Next
Dim LcCrit As String
LcCrit = "Select * from custo where CodigoProduto='" & GlCodigoProduto & "' order By DescricaoCusto"

Data1.DatabaseName = GLBase
Data1.RecordSource = LcCrit
Data1.Refresh
item.Refresh
End Sub

Private Sub Form_Load()
Dim Db          As Database
Dim RsCusto     As Recordset
Dim LcProduto   As String
Dim LcCrit      As String
Dim LcTotalC    As Long
Dim LcTotalPr   As Long
Dim LcBusca     As String
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2

LcCrit = "Select * from custo where CodigoProduto='" & GlCodigoProduto & "' order By DescricaoCusto"

Data1.DatabaseName = GLBase
Data1.RecordSource = LcCrit
Data1.Refresh
If Data1.Recordset.EOF Then
    Set Db = OpenDatabase(GLBase, False, False) ' "dBASE III;")
    Set RsCusto = Db.OpenRecordset("DecricaoCusto", dbOpenDynaset, dbSeeChanges, dbOptimistic)
    Do Until RsCusto.EOF
       Data1.Recordset.AddNew
       Data1.Recordset.Fields("CodigoDescricao") = RsCusto!Codigo
       Data1.Recordset.Fields("DescricaoCusto") = RsCusto!Descricao
       Data1.Recordset.Fields("CodigoProduto") = GlCodigoProduto
       Data1.Recordset.Update
       RsCusto.MoveNext
    Loop
    Data1.Recordset.Close
    Data1.DatabaseName = GLBase
    Data1.RecordSource = LcCrit
    Data1.Refresh
    item.Refresh
Else
  '=== Verifica se Todos já Estão Cadastrados
    Set Db = OpenDatabase(GLBase, False, False) ' "dBASE III;")
    Set RsCusto = Db.OpenRecordset("DecricaoCusto", dbOpenDynaset, dbSeeChanges, dbOptimistic)
    '=== Verifica A quantidade de reg no custos
    RsCusto.MoveLast
    LcTotalC = RsCusto.RecordCount
    RsCusto.MoveFirst
    '=== Verifica total no produto
    Data1.Recordset.MoveLast
    LcTotalPr = Data1.Recordset.RecordCount
    Data1.Recordset.MoveFirst
    If LcTotalPr < LcTotalC Then
       Do Until RsCusto.EOF
          LcBusca = "CodigoDescricao='" & RsCusto!Codigo & "'"
          Data1.Recordset.FindFirst LcBusca
          If Data1.Recordset.NoMatch Then
             Data1.Recordset.AddNew
             Data1.Recordset.Fields("CodigoDescricao") = RsCusto!Codigo
             Data1.Recordset.Fields("DescricaoCusto") = RsCusto!Descricao
             Data1.Recordset.Fields("CodigoProduto") = GlCodigoProduto
             Data1.Recordset.Update
          End If
          RsCusto.MoveNext
      Loop
    End If
    Data1.Recordset.Close
    Data1.DatabaseName = GLBase
    Data1.RecordSource = LcCrit
    Data1.Refresh
    item.Refresh
End If
End Sub

Private Sub item_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub item_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 46 Then KeyAscii = 44
End Sub

Private Sub OKButton_Click()
On Error Resume Next
Dim LcTotal     As String
Dim Lcusto      As Double
Dim LcustoAnt   As Double
Dim LcVenda     As Double
Dim LPerc       As Double
Dim LcCusto     As Double
If GlDecimais = 0 Then GlDecimais = 2
LcCusto = CDbl(FrmProduto.valor(3).Text)
LcustoAnt = CDbl(FrmProduto.valor(3).Text)
'==Calcula o Valor do Perentual atual do Preco Minimo
If Len(Trim(FrmProduto.valor(2).Text)) = 0 Then FrmProduto.valor(2).Text = 0
If Len(Trim(FrmProduto.valor(6).Text)) = 0 Then FrmProduto.valor(6).Text = 0

LPerc = (CDbl(AcertaNumero(FrmProduto.valor(2).Text, GlDecimais)) - CDbl(AcertaNumero(FrmProduto.valor(6).Text, GlDecimais))) / CDbl(AcertaNumero(FrmProduto.valor(6).Text, GlDecimais))
Data1.Recordset.MoveFirst
Do Until Data1.Recordset.EOF
   If Len(LcTotal) > 0 Then
      LcTotal = LcTotal & "+"
   End If
   LcTotal = LcTotal & Data1.Recordset.Fields("valor")
   LcCusto = AcertaNumero(CStr(LcCusto + (LcCusto * (Data1.Recordset.Fields("valor") / 100))), GlDecimais)
   Data1.Recordset.MoveNext
Loop
FrmProduto.valor(5).Text = LcTotal
FrmProduto.valor(6).Text = AcertaNumero(CStr(LcCusto), GlDecimais)
'=== Calcula Valor Venda
LcVenda = AcertaNumero(CStr(LcCusto + (LcCusto * (CDbl(FrmProduto.valor(0).Text) / 100))), GlDecimais)
FrmProduto.valor(1).Text = LcVenda
'=== Calcula o Valor Minimo De Venda
LPerc = LPerc + 1
FrmProduto.valor(2).Text = AcertaNumero(CStr(LPerc * CDbl(FrmProduto.valor(6).Text)), GlDecimais)

If FrmProduto.valor(2).Text = "-1,#I" Then FrmProduto.valor(2).Text = 0

GlCampo8 = FrmProduto.valor(2).Text
GlCampo22 = FrmProduto.valor(4).Text
GlCampo14 = FrmProduto.valor(1).Text
GlCampo12 = FrmProduto.valor(3).Text
GlCampo17 = FrmProduto.valor(0).Text
GlCampo21 = FrmProduto.valor(5).Text
GlCampo23 = FrmProduto.valor(6).Text
FrmProduto.CmdSalvar.Enabled = True
Unload Me
End Sub

Private Sub OKButton_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub
