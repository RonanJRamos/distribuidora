VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form RelResumoProdutoCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Produtos Comprados Periodo"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6180
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   6180
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Codigo 
      Height          =   285
      Left            =   5040
      TabIndex        =   7
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar"
      Height          =   495
      Left            =   1560
      TabIndex        =   6
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Visualizar"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   1335
   End
   Begin Crystal.CrystalReport CryRelatorio 
      Left            =   3000
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSMask.MaskEdBox Dataf 
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Datai 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin VB.ComboBox Cliente 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "RelResumoProdutoCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type MtCliente
    codigo As String
    Nome As String
End Type
Private Mt() As MtCliente
Sub carregaCliente()
Dim Rs As Recordset
Dim a As Integer

AbreBase

Set Rs = Dbbase.OpenRecordset("Select * from alid001 order by razaosoc")
Do Until Rs.EOF
  ReDim Preserve Mt(a)
  Mt(a).codigo = Rs!codigo
  Mt(a).Nome = Rs!razaosoc & ""
  cliente.AddItem Rs!razaosoc & ""
  a = a + 1
  Rs.MoveNext
Loop
Set Rs = Nothing


End Sub

Private Sub Cliente_Click()
On Error Resume Next
codigo.Text = Mt(cliente.ListIndex).codigo
End Sub

Private Sub Command1_Click()
Dim StrSql As String
Dim Rs As ADODB.Recordset
Dim RsItens As ADODB.Recordset
Dim RsRel As Recordset
Dim TotalNota As Long
Dim a As Integer

On Error GoTo erroS
StrSql = "SELECT * from alid050 WHERE ((DTEMIS) Between '" & Format(Datai.Text, "yy-mm-dd") & "' And '" & Format(Dataf.Text, "yy-mm-dd") & "') and (Cliente='" & codigo.Text & "');"
Set Rs = AbreRecordset(StrSql)
Rs.MoveLast
TotalNota = Rs.RecordCount
Rs.MoveFirst
Debug.Print StrSql

Dbbase.Execute "delete from RelatorioProdutoComprado"

Set RsRel = Dbbase.OpenRecordset("Select * from RelatorioProdutoComprado")
a = 1
Do Until Rs.EOF
  Me.Caption = "Processando nota " & a & " de " & TotalNota
  DoEvents
  StrSql = "Select * from alid052 where NumNf='" & Rs!numnf & "'"
  Set RsItens = AbreRecordset(StrSql)
  Do Until RsItens.EOF
    RsRel.AddNew
    RsRel!DTEMIS = Rs!DTEMIS
    RsRel!cliente = Rs!cliente & ""
    RsRel!qtde = CDbl(RsItens!qtde) * CDbl(RsItens!QTDUM)
    RsRel!VALUNIT = RsItens!VALUNIT
    RsRel!Unimed = "UN"
    RsRel!QTDUM = 1
    RsRel!Descricao = RsItens!Descricao & ""
    RsRel!codProd = RsItens!codProd
    RsRel.Update
    RsItens.MoveNext
  Loop
  Set RsItens = Nothing
  Rs.MoveNext
  a = a + 1
Loop
Set Rs = Nothing
Set RsRel = Nothing
  
CryRelatorio.DataFiles(0) = GLBase
CryRelatorio.ReportFileName = App.Path & "\ProdutosClientePeriodo.rpt"

CryRelatorio.DiscardSavedData = True
CryRelatorio.WindowTop = 50
CryRelatorio.WindowWidth = 700
CryRelatorio.WindowLeft = 50
CryRelatorio.WindowHeight = 500
CryRelatorio.WindowTitle = "Compras de clientes por periodo"

LcTipoSaida = 0

CryRelatorio.Destination = LcTipoSaida
CryRelatorio.PrintReport

If CryRelatorio.LastErrorNumber > 0 Then MsgBox CryRelatorio.LastErrorString
Exit Sub
erroS:
MsgBox err.Description & err.Description

Resume Next

End Sub

Private Sub Command2_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub Form_Load()
carregaCliente
End Sub
