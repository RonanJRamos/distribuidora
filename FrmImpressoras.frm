VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FrmImpressoras 
   Caption         =   "Cadastro de Impressoras"
   ClientHeight    =   3705
   ClientLeft      =   3615
   ClientTop       =   2655
   ClientWidth     =   9705
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3705
   ScaleWidth      =   9705
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\Projeto\Lids\banco\lidis.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Impressoras"
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Fechar  F10"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Width           =   1935
   End
   Begin MSDBGrid.DBGrid impressora 
      Bindings        =   "FrmImpressoras.frx":0000
      Height          =   2775
      Left            =   120
      OleObjectBlob   =   "FrmImpressoras.frx":0014
      TabIndex        =   0
      Top             =   240
      Width           =   9495
   End
End
Attribute VB_Name = "FrmImpressoras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
FrmOpcoes.portaorcamento.Clear
FrmOpcoes.carregacombo
Unload Me
End Sub

Private Sub Form_Load()
Data1.DatabaseName = GLBase
Data1.Refresh
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
End Sub
