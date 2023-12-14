VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Tabelasxcluidas 
   Caption         =   "Exibe Dados Excluidos"
   ClientHeight    =   810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   ScaleHeight     =   810
   ScaleWidth      =   5835
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Restaurar Registro"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3600
      TabIndex        =   3
      Top             =   240
      Width           =   1935
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSDBGrid.DBGrid db 
      Bindings        =   "Tabelasxcluidas.frx":0000
      Height          =   735
      Left            =   120
      OleObjectBlob   =   "Tabelasxcluidas.frx":0014
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox Tabela 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Tabela"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "Tabelasxcluidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lcarq       As String
Private GlAlterar As Boolean
Private GlIncluir As Boolean
Private GlExcluir As Boolean
Private LcPrimeiro As Boolean
Private LcTb        As String
'=== Estas sao locais, nao copiar
Private Type Tacampod
    Tamanho As Integer
    Nome    As String
End Type
Private Type CamposD
    codigo As String
    Nome As String
End Type

Private Mt() As CamposD
Private MtTamanho() As Tacampod

Private Sub Command1_Click()
On Error Resume Next
Dim rs As Recordset
Dim campo As Field
AbreBase
Set rs = Dbbase.OpenRecordset("Select * from " & LcTb, dbOpenDynaset, dbSeeChanges, dbOptimistic)
rs.AddNew
err.Number = 0
For Each campo In Data1.Recordset.Fields
    If err.Number <> 0 Then Exit For
    LcCampo = campo.Name
    If Len(Data1.Recordset.Fields(LcCampo)) > 0 Then
       If LcCampo <> "maquinaExclusao" And LcCampo <> "usuarioExclusao" _
       And LcCampo <> "dataexclusao" And LcCampo <> "horaexclusao" And LcCampo <> "tabelaexcluida" Then
            rs.Fields(LcCampo) = Data1.Recordset.Fields(LcCampo)
       End If
    End If
Next
rs.Update
Data1.Recordset.Delete
rs.Close
Dbbase.Close
Set rs = Nothing
Set Dbbase = Nothing
Tabela.SetFocus
Data1.Refresh
If Not Data1.Recordset.EOF Then
   If Me.WindowState <> 2 Then Me.WindowState = 2
   db.Visible = True
   Command1.Enabled = True
 Else
   Command1.Enabled = False
   Me.WindowState = 0
   'MsgBox "Não Foram Excluídos Registros da Tabela " & Tabela.Text, 64, "Aviso"
 End If

'data1.Refresh
End Sub

Private Sub db_ColResize(ByVal ColIndex As Integer, Cancel As Integer)

   MtTamanho(ColIndex).Tamanho = db.Columns(ColIndex).Width
   MtTamanho(ColIndex).Nome = db.Columns(ColIndex).Caption
   GravarAtualizacaoIni
End Sub

Private Sub db_LostFocus()
'Command1.Enabled = False
End Sub

Private Sub Form_Load()
Dim LcPAth As String
Dim a As Integer
CarregaCombo
For a = Len(GLBase) To 1 Step -1
    LCLEtra = Mid(GLBase, a, 1)
    If LCLEtra = "\" Then
       LcPAth = Mid(GLBase, 1, a)
       Exit For
    End If
Next
Lcarq = LcPAth & "gradeexclusao.ini"
LcPrimeiro = True
End Sub

Function CarregaCombo()

Tabela.AddItem "Clientes"
Tabela.AddItem "Fornecedores"
Tabela.AddItem "Produtos"
Tabela.AddItem "Funcionarios"
Tabela.AddItem "Galpões"
Tabela.AddItem "Cidade"
Tabela.AddItem "Tipo Monetário"
Tabela.AddItem "Tipo de Receitas e Despesas"

Tabela.AddItem "Unidade"
Tabela.AddItem "Transportadora"
Tabela.AddItem "Custo"
Tabela.AddItem "Receitas"
Tabela.AddItem "Despesas"
Tabela.AddItem "Cheques Recebidos"
Tabela.AddItem "Notas de Entrada"
Tabela.AddItem "Dados da Notas de Entrada"
Tabela.AddItem "Notas de Saídas"
Tabela.AddItem "Dados da Notas de Saídas"
Tabela.AddItem "Pedidos de Vendas"
Tabela.AddItem "Dados do Pedido de Vendas"

Tabela.AddItem "Orçamento e Vendas"
Tabela.AddItem "Dados do Orçamento e Vendas"

End Function

Private Sub Form_Resize()
acertaTamanhogrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
GravarAtualizacaoIni
End Sub

Private Sub Tabela_Click()
On Error Resume Next
LcCap = Me.Caption
Me.Caption = "Aguarde, Procurando Registros Excluídos em " & Tabela.Text
DoEvents
Select Case Tabela.Text
    Case "Clientes"
        LcTb = "Alid001"
    Case "Fornecedores"
        LcTb = "Alid002"
    Case "Produtos"
        LcTb = "Alid009"
    Case "Funcionarios"
        LcTb = "Alid200"
    Case "Galpões"
        LcTb = "Alid012"
    Case "Cidade"
        LcTb = "Alid005"
    Case "Tipo Monetário"
        LcTb = "Alid008"
    Case "Tipo de Receitas e Despesas"
        LcTb = "Alid007"
    Case "Unidade"
        LcTb = "Alid004"
    Case "Transportadora"
        LcTb = "Transportadora"
    Case "Custo"
        LcTb = "DecricaoCusto"
    Case "Receitas"
        LcTb = "Alid015"
    Case "Despesas"
        LcTb = "Alid014"
    Case "Cheques Recebidos"
        LcTb = "cheques"
    Case "Notas de Entrada"
        LcTb = "EntradaNf"
    Case "Dados da Notas de Entrada"
        LcTb = "ItensEntradaNf"
    Case "Notas de Saídas"
        LcTb = "alid050"
    Case "Dados da Notas de Saídas"
        LcTb = "alid052"
    Case "Pedidos de Vendas"
        LcTb = "proposta"
    Case "Dados do Pedido de Vendas"
        LcTb = "subproposta"
    Case "Orçamento e Vendas"
        LcTb = "Orcamento"
    Case "Dados do Orçamento e Vendas"
        LcTb = "DadosOrcamento"
End Select

GlTitulo = "Dados da Tabela " & Tabela.Text
'GLBase = "c:\projeto\banco\lidis.mdb"
Me.Caption = GlTitulo
GlTabela = "Select * from deletada where tabelaexcluida='" & LcTb & "'"
GlAlterar = False
GlIncluir = False
GlExcluir = False

db.AllowUpdate = GlAlterar
db.AllowDelete = GlExcluir
db.AllowAddNew = GlIncluir

Data1.DatabaseName = GLBase
Data1.RecordSource = GlTabela
Data1.Refresh
AcertaTamanho
LcPrimeiro = False
Me.Caption = LcCap
If Not Data1.Recordset.EOF Then
   If Me.WindowState <> 2 Then Me.WindowState = 2
   db.Visible = True
   Command1.Enabled = True
 Else
   Command1.Enabled = False
   MsgBox "Não Foram Excluídos Registros da Tabela " & Tabela.Text, 64, "Aviso"
 End If
End Sub
Function acertaTamanhogrid()
On Error Resume Next
db.Height = Me.Height - 1400
db.Width = Me.Width - 100
db.Top = 850
db.Left = 0
Lct = Me.Width / 2
'StatusBar.Panels(1).Width = Lct
'StatusBar.Panels(2).Width = Lct
End Function
Function AcertaTamanho()
Dim a As Integer
ReDim MtTamanho(0)

For a = 0 To Data1.Recordset.Fields.Count - 1
    ReDim Preserve MtTamanho(a)
    
    Lct = LeIni(Tabela.Text, db.Columns(a).Caption, Lcarq)
    If Len(Lct) > 0 Then
       MtTamanho(a).Tamanho = Lct
    Else
       MtTamanho(a).Tamanho = db.Columns(a).Width
    End If
    If Len(db.Columns(a)) = 0 Then
       db.Columns(a).Visible = False
    Else
       db.Columns(a).Visible = True
    End If
       
    db.Columns(a).Width = MtTamanho(a).Tamanho

    MtTamanho(a).Nome = db.Columns(a).Caption
    If Data1.Recordset.Fields(MtTamanho(a).Nome).Type = dbDate Then db.Columns(a).NumberFormat = "Short Date"
    If Data1.Recordset.Fields(MtTamanho(a).Nome).Type = dbCurrency Then db.Columns(a).NumberFormat = "Fixed"
    If Data1.Recordset.Fields(MtTamanho(a).Nome).Type = dbDouble Then db.Columns(a).NumberFormat = "Fixed"
    If Data1.Recordset.Fields(MtTamanho(a).Nome).Type = dbNumeric Then db.Columns(a).NumberFormat = "Fixed"
         
Next

End Function
Function GravarAtualizacaoIni()

On Error Resume Next
Dim lcchave As String
Dim LcTamanh As String
Dim a As Integer

For a = 0 To Data1.Recordset.Fields.Count - 1
    lcchave = MtTamanho(a).Nome
    LcTamanh = CStr(MtTamanho(a).Tamanho)
    Call GravaIni(Tabela.Text, lcchave, LcTamanh, Lcarq)
Next

End Function

