VERSION 5.00
Begin VB.Form FrmEspecie 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Funções"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   10245
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt 
      Enabled         =   0   'False
      Height          =   405
      Index           =   0
      Left            =   1440
      TabIndex        =   8
      TabStop         =   0   'False
      Tag             =   "S/N/S/00/N/CODIGO"
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox txt 
      Height          =   405
      Index           =   1
      Left            =   1440
      MaxLength       =   70
      TabIndex        =   7
      Tag             =   "S/T/S/01/N/ESPECIE"
      Top             =   720
      Width           =   8655
   End
   Begin VB.CommandButton CmdFechar 
      Caption         =   "&Fechar"
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton CmdSalvar 
      Caption         =   "&Salvar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1455
      TabIndex        =   5
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton CmdPrimeiro 
      Caption         =   "&Primeiro"
      Height          =   375
      Left            =   4365
      TabIndex        =   4
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton CmdExcluir 
      Caption         =   "&Excluir"
      Height          =   375
      Left            =   2910
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton CmdAnterior 
      Caption         =   "&Anterior"
      Height          =   375
      Left            =   5820
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton CmdSeguinte 
      Caption         =   "&Seguinte"
      Height          =   375
      Left            =   7275
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton CmdUltimo 
      Caption         =   "&Ultimo"
      Height          =   375
      Left            =   8730
      TabIndex        =   0
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   10320
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   1020
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cargo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Width           =   570
   End
End
Attribute VB_Name = "FrmEspecie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdFechar_Click()
Unload Me
End Sub

Private Sub Form_Activate()
On Error Resume Next
If LcCarregado Then Exit Sub
Select Case LcTipoDados
   Case Is = 1
        DesabilitaCtr
   Case Is = 2
      Call AbreBanco(especie)
      VinculaDados
   Case Is = 3
      'DesabilitaTodos
      MnSalvar.Enabled = False
      MnExcluir.Enabled = False
      Call AbreBanco(especie)
      CmdExcluir.Enabled = False
      VinculaDados
 End Select
'CriaMascara
LcRegAtual = False
FrmPrincipal.Visible = False
CarreGamatriz
LcCarregado = True

End Sub
Private Function DesabilitaCtr()
CmdPrimeiro.Enabled = False
CmdAnterior.Enabled = False
CmdUltimo.Enabled = False
CmdSeguinte.Enabled = False
MnMovimento.Enabled = False
MnRegistro.Enabled = False
CmdExcluir.Enabled = False
End Function
Function VinculaDados()
On Error Resume Next


If LcTipoDados = 1 Then NovoReg Else Call RegistroAtual(especie)



txt(0).Text = GlCampo0
txt(1).Text = GlCampo1
End Function
Function CarreGamatriz()
Dim a As Integer, LcNome As String, LcTipo As String
GlFormAtual = Tabela.especie

Set GlFormA = Me
For a = 0 To 30
    LcNome = Mid$(txt(a).Tag, 12)
    LcTipo = Mid$(txt(a).Tag, 3, 1)
    If Err <> 0 Then Exit For
    MtPesquisa(a).Indice = LcNome
    MtPesquisa(a).tipo = LcTipo
    MtPesquisa(a).Campo = LcNome
    
 Next
 
End Function
Private Sub Form_Load()
On Error Resume Next

DataS.Text = Format(GlDataSistema, "dd/mm/yyyy")
HoraS.Text = Format(Time, "hh:mm:ss")
Top = 0
Left = Screen.Width / 2 - Width / 2
LcIndice = "CODIGO"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

If (LcTipoDados = 1) And (CmdSalvar.Enabled = True) Then
   GlPergunta = True
   SalvaRegistro (Cliente)
End If
If (LcTipoDados = 2) And LcAlterado Then SalvaRegistro (Cliente)
FechaBanco
FrmPrincipal.Visible = True
LcCarregado = False
End Sub

Private Sub Txt_Change(Index As Integer)
Call Alterado
End Sub

Private Sub txt_GotFocus(Index As Integer)
If VerificaDuplicado(Index) Then
   txt(Index).SetFocus
End If
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
Call MoveTecla(Index, KeyCode)
End Sub
Private Sub CmdAnterior_Click()
On Error Resume Next
GlMov = True
If MovImentacao(enAnterior, especie) Then VinculaDados
GlMov = False
LcRegAtual = False
End Sub

Private Sub CmdExcluir_Click()
On Error Resume Next
If Exclui(especie) = 1 Then
      VinculaDados
End If
End Sub


Private Sub CmdPrimeiro_Click()
On Error Resume Next
GlMov = True
If MovImentacao(enPrimeiro, especie) Then VinculaDados
GlMov = False
LcRegAtual = False
End Sub

Private Sub CmdSalvar_Click()
On Error Resume Next
Call SalvaRegistro(especie)
LcRegAtual = True
VinculaDados
LcRegAtual = False

End Sub

Private Sub CmdSeguinte_Click()
On Error Resume Next
GlMov = True

If MovImentacao(enSeguinte, especie) Then VinculaDados
GlMov = False
LcRegAtual = False
End Sub

Private Sub CmdUltimo_Click()
On Error Resume Next
GlMov = True
If MovImentacao(enultimo, especie) Then VinculaDados
GlMov = False
LcRegAtual = False
End Sub
