VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form ExibeCidade 
   BackColor       =   &H00D8C5B6&
   Caption         =   "Exibe Cidades Cadastradas"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5565
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   5565
   StartUpPosition =   3  'Windows Default
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "ExibeCidade.frx":0000
      Height          =   3015
      Left            =   120
      OleObjectBlob   =   "ExibeCidade.frx":0014
      TabIndex        =   5
      Top             =   1080
      Width           =   5295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Confirma F2"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ALID005"
      Top             =   120
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox criterio 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar  F10"
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
      Left            =   240
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4200
      Width           =   2055
   End
   Begin MSFlexGridLib.MSFlexGrid MostraCliente 
      Height          =   255
      Left            =   3840
      TabIndex        =   3
      Top             =   3480
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      FixedCols       =   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Criterio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   765
   End
End
Attribute VB_Name = "ExibeCidade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private a As Integer
Private Sub cidade_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 113 Then SendKeys "%+{C}"
End Sub

Private Sub Command1_Click()
On Error Resume Next
GlFormA.SetFocus
Select Case GlFormA.Name
   Case Is = "FrmCliente"
       FrmCliente.Txt(7).Text = Data1.Recordset.Fields(0)
       FrmCliente.Cidade.Caption = Data1.Recordset.Fields(1)
   Case Is = "FrmFornecedor"
       FrmFornecedor.Txt(7).Text = Data1.Recordset.Fields(0)
       FrmFornecedor.Cidade.Caption = Data1.Recordset.Fields(1)
  Case Is = "FrmGalpao"
       FrmGalpao.Txt(7).Text = Data1.Recordset.Fields(0)
       FrmGalpao.Cidade.Caption = Data1.Recordset.Fields(1)
 Case Is = "FrmTransportadora"
       FrmTransportadora.Txt(7).Text = Data1.Recordset.Fields(0)
       FrmTransportadora.Cidade.Caption = Data1.Recordset.Fields(1)
End Select
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

Private Sub criterio_Change()
On Error Resume Next

LcCriterio = "nome like '" & criterio.Text & "*'"
Data1.Recordset.FindFirst LcCriterio
End Sub

Private Sub criterio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 113 Then SendKeys "%+{C}"
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 121 Then SendKeys "%+{F}"
If KeyCode = 113 Then SendKeys "%+{C}"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 121 Then Command2_Click

End Sub

Private Sub Form_Load()
Data1.DatabaseName = GLBase
Data1.Refresh
Me.Top = 800
Me.Left = 6180
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
GlFormA.SetFocus
End Sub
