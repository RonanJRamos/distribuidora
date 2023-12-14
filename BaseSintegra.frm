VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form BaseSintegra 
   Caption         =   "Form1"
   ClientHeight    =   9210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10665
   LinkTopic       =   "Form1"
   ScaleHeight     =   9210
   ScaleWidth      =   10665
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdFiltrar 
      Caption         =   "Filtrar"
      Height          =   495
      Left            =   6000
      TabIndex        =   3
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox filtro 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   5655
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "BaseSintegra.frx":0000
      Height          =   8175
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   14420
      _Version        =   393216
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "CodigoProduto"
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
      BeginProperty Column01 
         DataField       =   "descricao"
         Caption         =   "Nome"
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
         DataField       =   "quantidade"
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   6614,93
         EndProperty
         BeginProperty Column02 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Filrar produtos começados com"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   2205
   End
End
Attribute VB_Name = "BaseSintegra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdFiltrar_Click()
Dim Rs As ADODB.Recordset
Dim StrSql As String
StrSql = "Select * from basesintegra  where Descricao like '" & filtro.Text & "%' order by descricao"
Set Rs = AbreRecordset(StrSql)
Set DataGrid1.DataSource = Rs


End Sub

Private Sub Form_Load()
'On Error Resume Next
'dados.ConnectionString = conexaoAdo.ConnectionString
'dados.RecordSource = "Select * from basesintegra order by descricao"
'If dados.Recordset.EOF Then buscaprodutos
Dim Rs As ADODB.Recordset
Dim StrSql As String
StrSql = "Select * from basesintegra order by descricao"
Set Rs = AbreRecordset(StrSql)
Set DataGrid1.DataSource = Rs
If Rs.EOF Then buscaprodutos
End Sub
Sub buscaprodutos()
Dim StrSql As String
Dim Rs As ADODB.Recordset
StrSql = "Select * from produtos order by nome"
Set Rs = AbreRecordset(StrSql, True)

Do Until Rs.EOF
  Me.Caption = "Filtrado produto " & Rs!Nome
  
    DoEvents
  StrSql = "Insert into basesintegra (codigoproduto,descricao,quantidade) values ("
  StrSql = StrSql & Rs!Codigo & ",'"
  StrSql = StrSql & Replace(Rs!Nome, "'", "''") & "',0)"
  ExecutaSql StrSql
  
  Rs.MoveNext
Loop

StrSql = "Select * from basesintegra order by descricao"
Set Rs = AbreRecordset(StrSql)
Set DataGrid1.DataSource = Rs
If Rs.EOF Then buscaprodutos
End Sub
Private Sub Label2_Click()

End Sub

Private Sub Text1_Change()

End Sub
