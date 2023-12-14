VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Despesasnaoquitadas 
   BackColor       =   &H00E4E3D6&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Despesas Não Quitadas"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10425
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   10425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdFiltrar 
      Caption         =   "Filtrar"
      Height          =   375
      Left            =   7080
      TabIndex        =   11
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox Doc 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E4E3D6&
      Caption         =   "Ordenar Por"
      Height          =   735
      Left            =   3720
      TabIndex        =   3
      Top             =   240
      Width           =   3255
      Begin VB.OptionButton OptVencimento 
         BackColor       =   &H00E4E3D6&
         Caption         =   "vencimento"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton OptDoc 
         BackColor       =   &H00E4E3D6&
         Caption         =   "Documento"
         Height          =   195
         Left            =   1680
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
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
      Top             =   6600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar F10"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   6600
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Confirma F2"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   6600
      Width           =   2535
   End
   Begin MSDBGrid.DBGrid Desp 
      Bindings        =   "Despesasnaoquitadas.frx":0000
      Height          =   5175
      Left            =   120
      OleObjectBlob   =   "Despesasnaoquitadas.frx":0014
      TabIndex        =   0
      Top             =   1200
      Width           =   9975
   End
   Begin MSMask.MaskEdBox DataI 
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   600
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "99/99/99"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox DataF 
      Height          =   255
      Left            =   2640
      TabIndex        =   8
      Top             =   600
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "99/99/99"
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E4E3D6&
      Caption         =   "Documento"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E4E3D6&
      Caption         =   "Periodo Vencimento"
      Height          =   255
      Left            =   1560
      TabIndex        =   9
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "Despesasnaoquitadas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdFiltrar_Click()
On Error Resume Next

Dim LcCriterio As String

LcCriterio = "Select * from alid014 where   nf like '" & Doc.Text & "*' and ((((ALID014.VALPAGO)=0)) OR (((ALID014.VALPAGO) Is Null)))" ' order by DTVENC"

If IsDate(DataI.Text) And IsDate(DataF.Text) Then
   LcCriterio = LcCriterio & " and DTVENC between #" & Format(DataI.Text, "mm/dd/yy") & "# and #" & Format(DataF.Text, "mm/dd/yy") & "#"
Else
  If IsDate(DataI.Text) Then
    LcCriterio = LcCriterio & " and DTVENC = #" & Format(DataI.Text, "mm/dd/yy") & "#"
  End If
  
End If
If OptVencimento.Value Then
    LcCriterio = LcCriterio & " order by DTVENC"
End If
If OptDoc.Value Then
    LcCriterio = LcCriterio & " order by nf"
End If
Debug.Print LcCriterio
Data1.DatabaseName = GLBase
Data1.RecordSource = LcCriterio
Data1.Refresh
End Sub

Private Sub Command1_Click()
On Error Resume Next
Dim LcCodigo    As String

LcCodigo = Data1.Recordset.Fields(1).Value
FrmBaixaDespesas.Txt(0).Text = LcCodigo
FrmBaixaDespesas.Codigo.Text = Data1.Recordset.Fields("Codigo").Value

Unload Me
FrmBaixaDespesas.Txt(0).SetFocus
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Desp_DblClick()
SendKeys "%{C}"
End Sub

Private Sub Desp_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   SendKeys "%{C}"
End If
If KeyCode = 121 Then Unload Me
If KeyCode = 113 Then SendKeys "%{C}"
End Sub

Private Sub Form_Load()
On Error Resume Next

Dim LcCriterio As String

LcCriterio = "Select * from alid014 where  (((ALID014.VALPAGO)=0)) OR (((ALID014.VALPAGO) Is Null)) order by DTVENC"
 Data1.DatabaseName = GLBase
 Data1.RecordSource = LcCriterio
 Data1.Refresh
End Sub
