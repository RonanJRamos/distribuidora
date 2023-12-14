VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form exibeDespRec 
   BackColor       =   &H00D8C5B6&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleciona Despesa - Receitas"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   ClipControls    =   0   'False
   Icon            =   "exibeDespRec.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5640
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox nome 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar F10"
      Height          =   615
      Left            =   2040
      TabIndex        =   2
      Top             =   5520
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Confirma F2"
      Default         =   -1  'True
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   5520
      Width           =   1695
   End
   Begin MSDBGrid.DBGrid item 
      Bindings        =   "exibeDespRec.frx":0442
      Height          =   4335
      Left            =   240
      OleObjectBlob   =   "exibeDespRec.frx":0456
      TabIndex        =   0
      Top             =   1080
      Width           =   5535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nome"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "exibeDespRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private a As Integer
Private Sub Command1_Click()
If GlFormA.Name = "Despesas" Then
   Despesas.Txt(11).Text = Data1.Recordset!cod & ""
   Despesas.Txt(12).Text = Data1.Recordset!Nome & ""
   Despesas.CmdSalvar.Enabled = True
   GlCampo11 = Data1.Recordset!cod & ""
   GlCampo12 = Data1.Recordset!Nome & ""
Else
   Receitas.Txt(11).Text = Data1.Recordset!cod & ""
   Receitas.Txt(12).Text = Data1.Recordset!Nome & ""
   Receitas.CmdSalvar.Enabled = True
   GlCampo11 = Data1.Recordset!cod & ""
   GlCampo12 = Data1.Recordset!Nome & ""
End If
Unload Me
End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 113 Then SendKeys "%+{C}"
End Sub

Private Sub Command2_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 113 Then SendKeys "%+{C}"
End Sub

Private Sub Form_Load()
On Error Resume Next
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
If GlFormA.Name = "Despesas" Then
   LcCrit = "Select * from alid007 where rd='D' order By Nome"
Else
  LcCrit = "Select * from alid007 where rd='R' order By Nome"
End If
Data1.DatabaseName = GLBase
Data1.RecordSource = LcCrit
Data1.Refresh
End Sub

Private Sub item_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 113 Then SendKeys "%+{C}"
End Sub

Private Sub nome_Change()
On Error Resume Next
LcCri = " nome like '" & Nome.Text & "'"
Data1.Recordset.FindFirst LcCri
End Sub

Private Sub Nome_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 113 Then SendKeys "%+{C}"
End Sub
