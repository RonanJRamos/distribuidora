VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form UltimasComprasClienteLocal 
   BackColor       =   &H00E6E4D2&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exibe as Últimas Compras do Cliente"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   11910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Confirmar"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Left            =   8160
      TabIndex        =   7
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox tipo 
      Height          =   375
      Left            =   8160
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox codigo 
      Height          =   285
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   360
      Width           =   855
   End
   Begin VB.ComboBox cliente 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   6975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Ver Contas Cliente F3"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   5400
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Fechar F10"
      Height          =   615
      Left            =   3120
      TabIndex        =   0
      Top             =   5400
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
      Left            =   5280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5400
      Visible         =   0   'False
      Width           =   1740
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4215
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   7435
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
      AutoSize        =   -1  'True
      BackColor       =   &H00E6E4D2&
      Caption         =   "Codigo"
      Height          =   195
      Index           =   1
      Left            =   7080
      TabIndex        =   8
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00E6E4D2&
      Caption         =   "Cliente"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "UltimasComprasClienteLocal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type DadoCliente
        Codigo As String
        Nome As String
End Type
Private MtCliente() As DadoCliente

Private LcTam, a As Long

Function carregaCliente()
Dim LcEmpresa, LcEndereco, LcFone, LcVer, LcCap, LcVer1 As String
Dim RsEmpresa As Recordset
AbreBase
LcTam = 0
Set RsEmpresa = Dbbase.OpenRecordset("Select * from alid001 order by RazaoSoc", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Do Until RsEmpresa.EOF
    ReDim Preserve MtCliente(LcTam)
    If Not IsNull(RsEmpresa!RAZAOSOC) Then
        MtCliente(LcTam).Codigo = RsEmpresa!Codigo
        MtCliente(LcTam).Nome = RsEmpresa!RAZAOSOC
        Cliente.AddItem RsEmpresa!RAZAOSOC
        LcTam = LcTam + 1
    End If
    RsEmpresa.MoveNext
Loop
If LcTam > 0 Then LcTam = LcTam - 1
RsEmpresa.Close
Dbbase.Close
Set RsEmpresa = Nothing
Set dbbasee = Nothing

End Function

Private Sub Cliente_Click()
On Error Resume Next
Codigo.Text = MtCliente(Cliente.ListIndex).Codigo
Command3.Enabled = Len(Codigo.Text)
'For a = 0 To LcTam
'    If MtCliente(a).Nome = cliente.Text Then
'       codigo.Text = MtCliente(a).codigo
'       Exit For
'    End If
'Next

End Sub

Private Sub Cliente_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{V}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

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
ContasClienteLocal.Show , Me
End Sub

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{V}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Command3_Click()
Dim LcSql As String
Dim RsNota As ADODB.Recordset
'AbreBase
LcSql = "select * from alid050 where CLIENTE='" & Codigo.Text & "' and DTEMIS>='2016-07-04' order by DTEMIS DESC"
'abreconexao
Set RsNota = AbreRecordsetRel(LcSql, RsNota)

If Not RsNota.EOF Then
   DataGrid1.Columns(7).DataField = "Status"
   'LcSql = "select * from alid050 where CLIENTE='" & codigo.Text & "' order by DTEMIS DESC"
   Tipo.Text = "N"
Else
  Tipo.Text = "O"
  DataGrid1.Columns(0).Caption = "Doc."
  DataGrid1.Columns(0).DataField = "Doc"
  DataGrid1.Columns(3).Caption = "V.Doc."
  DataGrid1.Columns(3).DataField = "TotalGeral"
  DataGrid1.Columns(7).DataField = "Status"
  LcSql = "select * from orcamento where CLIENTE='" & Codigo.Text & "' order by DTEMIS DESC"
End If
'MsgBox LcSql
'RsNota.Close
'Dbbase.Close

'Set rsa = AbreRecordset(LcSql, rsa)
Set DataGrid1.DataSource = RsNota

End Sub

Private Sub DataGrid1_DblClick()
On Error Resume Next
Me.Tag = DataGrid1.Columns(0)
detalhanotaLocal.Show , Me

End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then
    Me.Tag = DataGrid1.Columns(0)
    detalhanotaLocal.Show , Me
End If
End Sub

Private Sub Form_Load()
'On Error Resume Next
Dim Col1 As Column
'Set Col1 = Item.Columns(0)
'abreconexao
carregaCliente

Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
'FechaConexao
GlFormA.SetFocus
End Sub

Private Sub Item_DblClick()
On Error Resume Next
Me.Tag = Data1.Recordset.Fields(0)
detalhanotaLocal.Show , Me
End Sub

Private Sub item_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then
  Me.Tag = Data1.Recordset.Fields(0)
  detalhanotaLocal.Show , Me
End If
If KeyCode = 114 Then SendKeys "%+{V}"
If KeyCode = 121 Then SendKeys "%+{F}"
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

