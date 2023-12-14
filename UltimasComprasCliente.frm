VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form UltimasComprasCliente 
   BackColor       =   &H00E6E4D2&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exibe as Últimas Compras do Cliente"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3375
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   5953
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16772283
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
      ColumnCount     =   10
      BeginProperty Column00 
         DataField       =   "numnf"
         Caption         =   "Nota Fiscal"
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
         DataField       =   "dtemis"
         Caption         =   "Emissão"
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
         DataField       =   "natureza"
         Caption         =   "Natureza"
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
         DataField       =   "valornota"
         Caption         =   "Valor Nf"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """R$ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "vencimento1"
         Caption         =   "Vencimento 1"
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
         DataField       =   "vencimento2"
         Caption         =   "vencimento 2"
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
         DataField       =   "vencimento3"
         Caption         =   "Vencimento 3"
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
      BeginProperty Column07 
         DataField       =   "vencimento4"
         Caption         =   "Vencimento 4"
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
      BeginProperty Column08 
         DataField       =   "vencimento5"
         Caption         =   "Vencimento 5"
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
      BeginProperty Column09 
         DataField       =   ""
         Caption         =   "Status"
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
            ColumnWidth     =   1094,74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1019,906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1049,953
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1049,953
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1140,095
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1094,74
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1065,26
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1140,095
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1140,095
         EndProperty
         BeginProperty Column09 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Ver Contas Cliente  F3"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   7680
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Fechar F10"
      Height          =   615
      Left            =   3120
      TabIndex        =   0
      Top             =   7680
      Width           =   1575
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7560
      Visible         =   0   'False
      Width           =   1740
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   3375
      Left            =   120
      TabIndex        =   3
      Top             =   4080
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   5953
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16772283
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
      ColumnCount     =   10
      BeginProperty Column00 
         DataField       =   "numnf"
         Caption         =   "Nota Fiscal"
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
         DataField       =   "dtemis"
         Caption         =   "Emissão"
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
         DataField       =   "natureza"
         Caption         =   "Natureza"
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
         DataField       =   "valornota"
         Caption         =   "Valor Nf"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """R$ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "vencimento1"
         Caption         =   "Vencimento 1"
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
         DataField       =   "vencimento2"
         Caption         =   "vencimento 2"
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
         DataField       =   "vencimento3"
         Caption         =   "Vencimento 3"
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
      BeginProperty Column07 
         DataField       =   "vencimento4"
         Caption         =   "Vencimento 4"
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
      BeginProperty Column08 
         DataField       =   "vencimento5"
         Caption         =   "Vencimento 5"
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
      BeginProperty Column09 
         DataField       =   ""
         Caption         =   "Status"
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
            ColumnWidth     =   1094,74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1019,906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1049,953
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1049,953
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1140,095
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1094,74
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1065,26
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1140,095
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1140,095
         EndProperty
         BeginProperty Column09 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "AL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3600
      Width           =   2655
   End
End
Attribute VB_Name = "UltimasComprasCliente"
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
If KeyCode = 114 Then SendKeys "%+{V}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Command2_Click()
On Error Resume Next
ContasCliente.Show , Me
End Sub

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{V}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub DataGrid1_DblClick()
On Error Resume Next
Me.Tag = DataGrid1.Columns(0)
detalhanota.Show , Me

End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
a = DataGrid1.Col
If KeyCode = 13 Then
  Me.Tag = DataGrid1.Columns(0)
  detalhanota.Show , Me
End If
If KeyCode = 114 Then SendKeys "%+{V}"
If KeyCode = 121 Then SendKeys "%+{F}"

End Sub

Private Sub Form_Load()
'On Error Resume Next
Dim Col1 As Column
Dim Rsa As ADODB.Recordset
Dim RsAl As ADODB.Recordset

'Set Col1 = Item.Columns(0)
Dim StrSql As String
Dim LcSql As String
If GlFormA.Name <> "Orcamento" Then
   DataGrid1.Columns(7).DataField = "Status"
   LcSql = "select * from alid050 where CLIENTE='" & GlFormA.Txt(8).Text & "' and transmitida<>0 and DTEMIS>='2016-07-04' order by DTEMIS DESC"
   StrSql = "select * from saidas where CLIENTE='" & GlFormA.Txt(8).Text & "' order by DTEMIS DESC"
Else
   DataGrid1.Columns(0).Caption = "Doc."
   DataGrid1.Columns(0).DataField = "Doc"
   DataGrid1.Columns(3).Caption = "V.Doc."
   DataGrid1.Columns(3).DataField = "TotalGeral"
   DataGrid1.Columns(7).DataField = "Status"
   LcSql = "select * from orcamento where CLIENTE='" & Orcamento.CodigoCliente.Text & "' order by DTEMIS DESC"

End If
'abreconexao
Set RsAl = AbreRecordsetRel(StrSql, RsAl)
Set Rsa = AbreRecordsetRel(LcSql, Rsa)
Set DataGrid1.DataSource = Rsa
Set DataGrid2.DataSource = RsAl

'Top = Screen.Height / 2 - Height / 2
'Left = Screen.Width / 2 - Width / 2

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

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
'FechaConexao
GlFormA.SetFocus
End Sub

Private Sub Item_DblClick()
On Error Resume Next
Me.Tag = Data1.Recordset.Fields(0)
detalhanota.Show , Me
End Sub

Private Sub item_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then
  Me.Tag = Data1.Recordset.Fields(0)
  detalhanota.Show , Me
End If
If KeyCode = 114 Then SendKeys "%+{V}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub
