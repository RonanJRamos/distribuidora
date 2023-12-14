VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form detalhanotaLocal 
   BackColor       =   &H00E6E4D2&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exibe os Itens da Nota Fiscal"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11400
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Mostrar Todos Pedidos do Cliente F2"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   5280
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Fechar F10"
      Default         =   -1  'True
      Height          =   495
      Left            =   3480
      TabIndex        =   0
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   5520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5280
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5055
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   8916
      _Version        =   393216
      BackColor       =   16711379
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "item"
         Caption         =   "Item"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "codprod"
         Caption         =   "Codigo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "descricao"
         Caption         =   "Descrição do Produto"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "unimed"
         Caption         =   "Embalagem"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "qtdum"
         Caption         =   "c/"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "qtde"
         Caption         =   "Quantidade"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "valunit"
         Caption         =   "V.Unitario"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   750,047
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1319,811
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   4199,811
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   840,189
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   794,835
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1019,906
         EndProperty
         BeginProperty Column06 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "detalhanotaLocal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private a As Integer
Private Sub Command1_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%{M}"
If KeyCode = 121 Then SendKeys "%{F}"
End Sub
Function AbreRecordsetRel(LcSql As String, RsAtual As ADODB.Recordset) As ADODB.Recordset

On Error GoTo ErroAbreRs
LcComentario = "- AbreRecordset - Criando Nova Instancia do RecordSet."
Set RsAtual = New ADODB.Recordset
LcComentario = "- AbreRecordset - Setando os Parametros do Recordset."
RsAtual.CursorType = adOpenDynamic ' adOpenStatic
RsAtual.CursorLocation = adUseClient
RsAtual.LockType = adLockReadOnly
RsAtual.Source = LcSql
RsAtual.ActiveConnection = conexaoAdo

LcComentario = "- AbreRecordset - Abrindo o Recordset."
RsAtual.Open
Set AbreRecordsetRel = RsAtual
Exit Function

ErroAbreRs:
'If err.Number = 3709 Then
'   'abreconexao
'   Resume 0
'End If
'If LcExibemsg Then ErrosSistema = MsgBox(msg, 64, "erro Abrindo Tabela. ") Else ErrosSistema = 0
'MsgBox err.Description & err.Number
'Resume 0
logErro err.Number, err.Description, LcComentario
Resume Next
End Function

Private Sub Command2_Click()
On Error Resume Next
Dim LcCl As String
Dim Rsa As ADODB.Recordset

If UltimasComprasClienteLocal.Tipo.Text <> "O" Then
    DataGrid1.Columns(0).Caption = "NF"
    DataGrid1.Columns(0).DataField = "alid052.numnf"
    LcSql = "SELECT alid050.CLiente, alid052.codProd, alid052.QTDE, alid052.descricao, alid052.UNIMED, alid052.VALUNIT, alid052.numnf "
    LcSql = LcSql & "FROM alid052 INNER JOIN alid050 ON alid052.numnf = alid050.numnf "
    LcSql = LcSql & "WHERE alid050.CLiente='" & UltimasComprasClienteLocal.codigo.Text & "' order by  alid052.numnf "
Else
    LcCl = UltimasComprasClienteLocal.codigo.Text
   DataGrid1.Columns(0).Caption = "Doc"
   DataGrid1.Columns(0).DataField = "dadosorcamento.doc"
   DataGrid1.Columns(1).Caption = "Código"
   DataGrid1.Columns(1).DataField = "CodigoProduto"
   DataGrid1.Columns(3).DataField = "unid"
   DataGrid1.Columns(4).DataField = "com"
   DataGrid1.Columns(5).DataField = "quant"
   DataGrid1.Columns(6).DataField = "Unit"
   

    LcSql = "SELECT * "
    LcSql = LcSql & "FROM DadosOrcamento  INNER JOIN Orcamento ON DadosOrcamento.doc = Orcamento.doc "
    LcSql = LcSql & "WHERE Orcamento.CLiente='" & LcCl & "' order by dadosorcamento.doc asc"
    'MsgBox LcSql
End If

Set Rsa = AbreRecordset(LcSql)
Set DataGrid1.DataSource = Rsa

End Sub

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%{M}"
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%{M}"
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub Form_Load()
On Error Resume Next
'abreconexao
Dim Rsa As ADODB.Recordset
Dim LcSql As String
If UltimasComprasClienteLocal.Tipo.Text <> "O" Then
   '[CIDADE] & [ESTADO] AS Expr1
   LcSql = "select *  from alid052 where NUMNF='" & UltimasComprasClienteLocal.Tag & "' order by ITEM"

Else
   
   DataGrid1.Columns(1).Caption = "Código"
   DataGrid1.Columns(1).DataField = "CodigoProd"
   DataGrid1.Columns(3).DataField = "unid"
   DataGrid1.Columns(4).DataField = "com"
   DataGrid1.Columns(5).DataField = "quant"
   DataGrid1.Columns(6).DataField = "Unit"
   
   
   LcSql = "select * from DadosOrcamento where doc='" & UltimasComprasClienteLocal.Tag & "' order by ITEM"
End If

Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2

Set Rsa = AbreRecordsetRel(LcSql, Rsa)
Set DataGrid1.DataSource = Rsa

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
'FechaConexao
End Sub
