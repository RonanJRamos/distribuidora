VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form FrmRelClienteTelemarketing 
   Caption         =   "Relatório de Clientes por Telemarketing"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6960
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6960
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox GrupoEconomico 
      Height          =   315
      ItemData        =   "FrmRelClienteTelemarketing.frx":0000
      Left            =   120
      List            =   "FrmRelClienteTelemarketing.frx":0002
      TabIndex        =   11
      Top             =   1440
      Visible         =   0   'False
      Width           =   4335
   End
   Begin Crystal.CrystalReport CryRelatorio 
      Left            =   4440
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.ComboBox Vendedor 
      Height          =   315
      ItemData        =   "FrmRelClienteTelemarketing.frx":0004
      Left            =   120
      List            =   "FrmRelClienteTelemarketing.frx":0006
      TabIndex        =   10
      Top             =   840
      Width           =   4335
   End
   Begin VB.TextBox codigo 
      Height          =   405
      Left            =   5280
      TabIndex        =   9
      Top             =   2760
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Copias 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   7
      Text            =   "1"
      Top             =   2280
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Saída"
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   1815
      Begin VB.OptionButton Video 
         Caption         =   "Vídeo"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Impressora 
         Caption         =   "Impressora"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar  F10"
      Height          =   495
      Left            =   5280
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Confirmar F3"
      Height          =   495
      Left            =   5280
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grupo Economico"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   5280
      TabIndex        =   8
      Top             =   2040
      Width           =   750
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Relatório de Clientes por Telemarketing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   6855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Telemarketing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1515
   End
End
Attribute VB_Name = "FrmRelClienteTelemarketing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type TipoVend
      codigo As String
      Nome As String
End Type
Private LcTamanho, a  As Integer
Private MtVendedor() As TipoVend

Private Sub Comissao_Change()

End Sub
Function CarregaGrupoEconomico()
On Error GoTo errc
Debug.Print conexaoAdo.ConnectionString
Dim RsVendedor As ADODB.Recordset
Set RsVendedor = AbreRecordset("Select * from GrupoEconomico order by nome", True)
LcTamanhoGr = 0
Do Until RsVendedor.EOF
   GrupoEconomico.AddItem RsVendedor!Nome
   RsVendedor.MoveNext
Loop
RsVendedor.Close
Set RsVendedor = Nothing
Exit Function
errc:
MsgBox err.Description & " " & err.Number
Exit Function

End Function
Function CarregaTelemarketing()
On Error GoTo errc
Dim RsVendedor As Recordset
AbreBase
Set RsVendedor = Dbbase.OpenRecordset("ALID200", dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcTamanho = 0
Do Until RsVendedor.EOF
   ReDim Preserve MtVendedor(LcTamanho)
   MtVendedor(LcTamanho).codigo = RsVendedor!codigo
   MtVendedor(LcTamanho).Nome = RsVendedor!Nome
   vendedor.AddItem RsVendedor!Nome
   RsVendedor.MoveNext
   LcTamanho = LcTamanho + 1
Loop
If LcTamanho > 0 Then LcTamanho = LcTamanho - 1
'Comissao.AddItem "TODOS"
'Comissao.Text = "TODOS"
RsVendedor.Close
Set RsVendedor = Nothing
Exit Function
errc:

Exit Function

End Function


Private Sub Command1_Click()
On Error Resume Next
Dim LcFormula, LcCriterio As String, LcTipoSaida As Integer
Dim RsEmpresa As Recordset, RsOpcao As Recordset
Dim LcEmpresa, LcEndereco, LcFone, LcCelular, Lccelular1, Lcemail, LcVer, LcCap, LcVer1 As String
AcertaUltimaCompra

AbreBase
Set RsEmpresa = Dbbase.OpenRecordset("Empresa", dbOpenDynaset, dbSeeChanges, dbOptimistic)
'Set RsOpcao = DbBase.OpenRecordset("Opcoes", dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcCap = Me.Caption
Me.Caption = "Aguarde, Gerando o Relatório..."
If Not RsEmpresa.EOF Then
   LcEmpresa = RsEmpresa!Razao
   LcEndereco = RsEmpresa!Endereco & " " & RsEmpresa!Bairro
   LcFone = RsEmpresa!Fone
End If
 
'Abertura do relatório de vendas
    
    
   CryRelatorio.DataFiles(0) = GLBase
   If GlImprimeSemLinha Then
      CryRelatorio.ReportFileName = App.Path & "\RelClienteTeleMarketing.rpt"
   Else
      CryRelatorio.ReportFileName = App.Path & "\RelClienteTeleMarketingsl.rpt"
   End If
    
    LcFormula = "UpperCase({ALID001.TelemarketingAtende})like '*" & UCase(vendedor.Text) & "*' and UpperCase({ALID001.GrupoEconomicoNome}) like '*" & UCase(GrupoEconomico.Text) & "*'"
    CryRelatorio.SortFields(0) = "+{ALID001.razaosoc}"
    CryRelatorio.CopiesToPrinter = Val(Copias.Text)

   
    
CryRelatorio.DiscardSavedData = True
CryRelatorio.WindowTop = 50
CryRelatorio.WindowWidth = 700
CryRelatorio.WindowLeft = 50
CryRelatorio.WindowHeight = 500
CryRelatorio.WindowTitle = "Clientes por Telemarkting"

CryRelatorio.Formulas(0) = "Empresa='" & LcEmpresa & "'"
CryRelatorio.Formulas(1) = "Endereco='" & LcEndereco & "'"
CryRelatorio.Formulas(2) = "Fone='" & LcFone & "'"
'CryRelatorio.Formulas(3) = "Versiculo='" & LcVer & "'"
'CryRelatorio.Formulas(4) = "Versiculo1='" & LcVer1 & "'"
CryRelatorio.Formulas(5) = "titulo='Clientes por Telemarketing'"
CryRelatorio.Formulas(3) = "Celular='" & LcCelular & "'"
CryRelatorio.Formulas(4) = "Celular1='" & Lccelular1 & "'"
CryRelatorio.Formulas(6) = "email='" & Lcemail & "'"
 If Impressora Then
   LcTipoSaida = 1
Else
   LcTipoSaida = 0
End If


CryRelatorio.SelectionFormula = LcFormula

CryRelatorio.Destination = LcTipoSaida
CryRelatorio.PrintReport
Me.Caption = LcCap
'RsOpcao.Close
RsEmpresa.Close
Dbbase.Close
Set RsOpcao = Nothing
Set RsEmpresa = Nothing
Set Dbbase = Nothing
If CryRelatorio.LastErrorNumber > 0 Then MsgBox CryRelatorio.LastErrorString

End Sub
Sub AcertaUltimaCompra()
On Error Resume Next
Dim LcCap As String
LcCap = Me.Caption
AbreBase
Dim rsCliente As Recordset
Dim RsNota As ADODB.Recordset
Dim StrSql As String
StrSql = "Select * from alid001 where TelemarketingAtende='" & vendedor.Text & "'"
Set rsCliente = Dbbase.OpenRecordset(StrSql, dbOpenDynaset, dbSeeChanges, dbOptimistic)
Dim TotalReg As Long
If Not rsCliente.EOF Then
    rsCliente.MoveLast
    TotalReg = rsCliente.RecordCount
    rsCliente.MoveFirst
End If
Do Until rsCliente.EOF
    a = a + 1
    Me.Caption = "Acertando ultima compra do Cliente " & rsCliente!razaosoc & " Reg:" & a & " de " & TotalReg
    DoEvents
    
   StrSql = "select dtemis,codigo from alid050 where status='Autorizado o uso da NF-e' and cliente='" & rsCliente!codigo & "' order by codigo desc limit 1"
   Set RsNota = AbreRecordset(StrSql, True)
   If Not RsNota.EOF Then
       StrSql = "Update alid001 set ULTCOMPRA=#" & Format(RsNota!DtEmis, "mm/dd/yy") & "# where codigo='" & rsCliente!codigo & "'"
       Dbbase.Execute StrSql
       x = Dbbase.RecordsAffected
   End If
   rsCliente.MoveNext
Loop
Me.Caption = LcCap
End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Copias_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Form_Load()
On Error Resume Next
DataS.Text = Format(GlDataSistema, "dd/mm/yyyy")
HoraS.Text = Format(Time, "hh:mm:ss")
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
LcIndice = "CODIGO"
Me.Height = 3495
Me.Width = 7080
CarregaTelemarketing
CarregaGrupoEconomico
End Sub

Private Sub Igual_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
FrmPrincipal.SetFocus
End Sub

Private Sub Impressora_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Iniciado_Click()
'Escolha
'BuscaExpressao

End Sub

Private Sub Iniciado_GotFocus()

End Sub

Private Sub Iniciado_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Nome_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Qualquer_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Vendedor_Change()
Dim a As Integer
For a = 0 To LcTamanho
    If MtVendedor(a).Nome = vendedor.Text Then
       codigo.Text = MtVendedor(a).codigo
       Exit For
    End If
Next
End Sub

Private Sub Vendedor_Click()
Dim a As Integer
For a = 0 To LcTamanho
    If MtVendedor(a).Nome = vendedor.Text Then
       codigo.Text = MtVendedor(a).codigo
       Exit For
    End If
Next
End Sub

Private Sub Vendedor_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Video_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub
