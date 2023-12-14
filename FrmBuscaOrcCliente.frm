VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FrmBuscaOrcCliente 
   BackColor       =   &H00CAE1A2&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busca Orçamennto / Vendas por Cliente"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7665
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5160
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Codigo 
      Height          =   375
      Left            =   5880
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Fechar F10"
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5160
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Confirma F2"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Left            =   240
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5160
      Width           =   2295
   End
   Begin VB.ComboBox cliente 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   5175
   End
   Begin MSDBGrid.DBGrid item 
      Bindings        =   "FrmBuscaOrcCliente.frx":0000
      Height          =   4095
      Left            =   240
      OleObjectBlob   =   "FrmBuscaOrcCliente.frx":0014
      TabIndex        =   1
      Top             =   960
      Width           =   7095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CLiente"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "FrmBuscaOrcCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Dacl
    Codigo As String
    Nome As String
End Type
Private MtCliente() As Dacl
Private LcTam, a As Long

Private Sub Cliente_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"

End Sub

Private Sub CLIENTE_LostFocus()
'On Error Resume Next
For a = 0 To LcTam - 1
    If Cliente.Text = MtCliente(a).Nome Then
       Codigo.Text = MtCliente(a).Codigo
       Exit For
    End If
Next
If Me.Tag = "proposta" Then
   Item.Columns(0).DataField = "numnf"
   LcSql = "select * from proposta where cliente='" & Codigo.Text & "' order by DTEMIS desc"
   
Else
   LcSql = "select * from orcamento where cliente='" & Codigo.Text & "' order by DTEMIS desc"
End If
Data1.DatabaseName = GLBase
Data1.RecordSource = LcSql
Data1.Refresh
Command1.Enabled = True
Item.SetFocus
End Sub

Private Sub Command1_Click()
On Error Resume Next
FrmPesquisaNota.Txt.Text = Data1.Recordset.Fields(0)
FrmPesquisaNota.cmdOK.SetFocus
Unload Me
End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Command3_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub Command3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Form_Load()
CarregaCli
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
End Sub
Function CarregaCli()
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

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
FrmPrincipal.SetFocus
End Sub

Private Sub item_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 114 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub
