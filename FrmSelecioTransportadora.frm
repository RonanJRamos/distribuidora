VERSION 5.00
Begin VB.Form FrmSelecioTransportadora 
   Caption         =   "Dados Transportadora"
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8565
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   8565
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Fechar F10"
      Height          =   495
      Left            =   3600
      TabIndex        =   5
      Top             =   4560
      Width           =   2895
   End
   Begin VB.CommandButton Confirma 
      Caption         =   "&Confirma F2"
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   4560
      Width           =   2895
   End
   Begin VB.ComboBox Transp 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   4455
   End
   Begin VB.TextBox Complemento 
      Height          =   2175
      Left            =   480
      MaxLength       =   250
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2160
      Width           =   7935
   End
   Begin VB.ComboBox Tipo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "FrmSelecioTransportadora.frx":0000
      Left            =   6840
      List            =   "FrmSelecioTransportadora.frx":000A
      TabIndex        =   1
      Text            =   "1- CIF"
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox fone 
      Height          =   285
      Left            =   1680
      MaxLength       =   30
      TabIndex        =   2
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transportadora "
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
      Index           =   0
      Left            =   0
      TabIndex        =   9
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
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
      Index           =   1
      Left            =   6240
      TabIndex        =   8
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Informações Complementares"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   9
      Left            =   2520
      TabIndex        =   7
      Top             =   1560
      Width           =   4185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fone"
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
      Index           =   11
      Left            =   120
      TabIndex        =   6
      Top             =   885
      Width           =   540
   End
End
Attribute VB_Name = "FrmSelecioTransportadora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type TipoTransp
    nome As String
    fone As String
End Type
Private TamMat, a As Long

Private MtTransp() As TipoTransp

Private Sub Command1_Click()
On Error Resume Next
Unload Me

End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 113 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Complemento_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 113 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Confirma_Click()
On Error Resume Next
orcamento.Transportadora.Text = Transp.Text & ""
orcamento.FoneTransp.Text = fone.Text & ""
orcamento.DadosComplementares.Text = Complemento.Text & ""

Unload Me
  
End Sub

Private Sub Confirma_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 113 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub fone_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Form_Load()
CarregaCombo
Transp.Text = orcamento.Transportadora.Text & ""
fone.Text = orcamento.FoneTransp.Text & ""
Complemento.Text = orcamento.DadosComplementares.Text & ""


End Sub

Function CarregaCombo()
Dim RsTransp As Recordset

AbreBase
Set RsTransp = Dbbase.OpenRecordset("Transportadora", dbOpenDynaset, dbSeeChanges, dbOptimistic)
TamMat = 0
Do Until RsTransp.EOF
   ReDim Preserve MtTransp(TamMat)
   MtTransp(TamMat).nome = RsTransp!Razaosoc & ""
   MtTransp(TamMat).fone = RsTransp!fone1 & ""
   Transp.AddItem RsTransp!Razaosoc & ""
   RsTransp.MoveNext
   TamMat = TamMat + 1
Loop
TamMat = TamMat - 1

RsTransp.Close
Set RsTransp = Nothing

   
End Function

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Unload Me
End Sub

Private Sub Tipo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Transp_Click()
For a = 0 To TamMat
   If Transp.Text = MtTransp(a).nome Then
      fone.Text = MtTransp(a).fone
      Exit For
   End If
Next
End Sub

Private Sub Transp_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{C}"
If KeyCode = 121 Then SendKeys "%+{F}"

End Sub
