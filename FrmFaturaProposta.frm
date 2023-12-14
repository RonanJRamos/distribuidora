VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "dbgrid32.ocx"
Begin VB.Form FrmFaturaProposta 
   Caption         =   "Propostas a Faturar"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   ScaleHeight     =   6315
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5880
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar F10"
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   5760
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Confirmar F2"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   5760
      Width           =   2055
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FrmFaturaProposta.frx":0000
      Height          =   5175
      Left            =   120
      OleObjectBlob   =   "FrmFaturaProposta.frx":0014
      TabIndex        =   0
      Top             =   480
      Width           =   11655
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Propostas a Faturar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "FrmFaturaProposta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private a As Integer
Private Sub Command1_Click()
On Error Resume Next
 Select Case GlFormA.Name
    Case Is = "FrmSaidaProdutoAlternativo"
        FrmSaidaProdutoAlternativo.proposta.Text = Data1.Recordset.Fields(0)
        FrmSaidaProdutoAlternativo.BuscaProposta (Data1.Recordset.Fields(0))
        Unload Me
        FrmSaidaProdutoAlternativo.SetFocus
   
    Case Is = "FrmSaidaProduto"
        FrmSaidaProduto.proposta.Text = Data1.Recordset.Fields(0)
        FrmSaidaProduto.BuscaProposta (Data1.Recordset.Fields(0))
        Unload Me
        FrmSaidaProduto.SetFocus
    
 End Select


End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Command2_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Form_Load()
Dim LcSql As String
On Error Resume Next
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
' item.Columns(0).Caption = "Doc."
DBGrid1.Columns(2).DataField = "razaosoc"
DBGrid1.Columns(4).DataField = "Nome"

LcSql = "SELECT proposta.NUMNF, proposta.DTEMIS, proposta.ValorNota, proposta.Previsao, proposta.Liberado, ALID001.RAZAOSOC, ALID200.NOME "
LcSql = LcSql & "FROM (proposta INNER JOIN ALID001 ON proposta.CLIENTE = ALID001.CODIGO) INNER JOIN ALID200 ON proposta.Vendedor = ALID200.CODIGO"
LcSql = LcSql & " where (Liberado=1) and (Previsao<=#" & Format(Date, "mm/dd/yy") & "#) and (faturado=False) and (pendente=False) order by proposta.DTEMIS desc"
'MsgBox LcSql

Data1.DatabaseName = GLBase
Data1.RecordSource = LcSql
Data1.Refresh
End Sub
