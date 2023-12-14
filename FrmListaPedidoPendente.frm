VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "dbgrid32.ocx"
Begin VB.Form FrmListaPedidoPendente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pedidos Pendentes"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   10050
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5760
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar"
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Confirma"
      Default         =   -1  'True
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   5520
      Width           =   1695
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FrmListaPedidoPendente.frx":0000
      Height          =   5055
      Left            =   240
      OleObjectBlob   =   "FrmListaPedidoPendente.frx":0014
      TabIndex        =   0
      Top             =   360
      Width           =   9615
   End
End
Attribute VB_Name = "FrmListaPedidoPendente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
FrmProposta.BuscaNota (Data2.Recordset.Fields("numnf"))
Unload Me
End Sub

Private Sub Command2_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
LcSql = "SELECT proposta.*, ALID001.RAZAOSOC, ALID200.NOME, proposta.Liberado, proposta.faturado, proposta.pendente "
LcSql = LcSql & "FROM (proposta INNER JOIN ALID001 ON proposta.CLIENTE = ALID001.CODIGO) INNER JOIN ALID200 ON proposta.Vendedor = ALID200.CODIGO "
LcSql = LcSql & "WHERE (((proposta.Liberado)=True) AND ((proposta.faturado)=False) AND ((proposta.pendente)=True));"
Debug.Print LcSql
Data2.DatabaseName = GLBase
Data2.RecordSource = LcSql
Data2.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
FrmProposta.SetFocus
End Sub
