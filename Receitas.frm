VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Receitas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Controle de Contas a Receber"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Tag             =   "GlCampo11 = RsAtual!codDesp & """""
   Begin VB.TextBox Txt 
      Height          =   375
      Index           =   12
      Left            =   2760
      TabIndex        =   10
      Top             =   4320
      Width           =   4815
   End
   Begin VB.TextBox Txt 
      Height          =   375
      Index           =   11
      Left            =   1320
      TabIndex        =   9
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox Txt 
      Height          =   1215
      Index           =   10
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   4920
      Width           =   6255
   End
   Begin VB.TextBox Txt 
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   5
      Top             =   2160
      Width           =   6255
   End
   Begin VB.CommandButton CmdFechar 
      Caption         =   "&Fechar F10"
      Height          =   615
      Left            =   8040
      TabIndex        =   20
      Top             =   4680
      Width           =   2385
   End
   Begin VB.CommandButton CmdSalvar 
      Caption         =   "&Salvar F2"
      Enabled         =   0   'False
      Height          =   615
      Left            =   8040
      TabIndex        =   13
      Top             =   1440
      Width           =   1185
   End
   Begin VB.CommandButton CmdExcluir 
      Caption         =   "&Excluir F3"
      Height          =   615
      Left            =   9240
      TabIndex        =   21
      Top             =   1440
      Width           =   1185
   End
   Begin VB.CommandButton CmdAnterior 
      Caption         =   "&Anterior F7"
      Height          =   615
      Left            =   9240
      TabIndex        =   17
      Top             =   3120
      Width           =   1185
   End
   Begin VB.CommandButton CmdSeguinte 
      Caption         =   "Se&guinte F8"
      Height          =   615
      Left            =   8040
      TabIndex        =   18
      Top             =   3840
      Width           =   1185
   End
   Begin VB.CommandButton CmdUltimo 
      Caption         =   "&Ultimo F9"
      Height          =   615
      Left            =   9240
      TabIndex        =   19
      Top             =   3840
      Width           =   1185
   End
   Begin VB.CommandButton CmdPrimeiro 
      Caption         =   "&Primeiro F6"
      Height          =   615
      Left            =   8040
      TabIndex        =   16
      Top             =   3120
      Width           =   1185
   End
   Begin VB.CommandButton CmdOrdenar 
      Caption         =   "&Ordenar F12"
      Height          =   615
      Left            =   9240
      TabIndex        =   15
      Top             =   2160
      Width           =   1185
   End
   Begin VB.CommandButton CmdPesquisar 
      Caption         =   "Pes&quisa F11"
      Height          =   615
      Left            =   8040
      TabIndex        =   14
      Top             =   2160
      Width           =   1185
   End
   Begin VB.TextBox Txt 
      Height          =   375
      Index           =   7
      Left            =   1320
      MaxLength       =   8
      TabIndex        =   12
      TabStop         =   0   'False
      Tag             =   "S/D/S/07/N/DTPAGTO"
      Top             =   6720
      Width           =   1095
   End
   Begin VB.TextBox Txt 
      Height          =   375
      Index           =   6
      Left            =   1320
      TabIndex        =   8
      Top             =   3480
      Width           =   6255
   End
   Begin MSMask.MaskEdBox Data 
      Height          =   375
      Index           =   0
      Left            =   1320
      TabIndex        =   1
      Top             =   1200
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yy"
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin VB.TextBox Txt 
      Enabled         =   0   'False
      Height          =   375
      Index           =   5
      Left            =   6000
      TabIndex        =   7
      Tag             =   "S/T/S/05/N/TPMONET"
      Top             =   480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Txt 
      Height          =   375
      Index           =   3
      Left            =   1320
      TabIndex        =   4
      Top             =   1680
      Width           =   6255
   End
   Begin VB.TextBox Txt 
      Enabled         =   0   'False
      Height          =   375
      Index           =   2
      Left            =   4680
      TabIndex        =   3
      Tag             =   "S/T/S/02/N/CLIENTE"
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Txt 
      Height          =   375
      Index           =   0
      Left            =   1320
      MaxLength       =   25
      TabIndex        =   0
      Tag             =   "S/T/S/00/N/Codigo"
      Top             =   600
      Width           =   4695
   End
   Begin VB.TextBox HoraS 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox DataS 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   360
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4560
      Top             =   0
   End
   Begin MSMask.MaskEdBox Data 
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   6
      Top             =   2880
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yy"
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Data 
      Height          =   375
      Index           =   2
      Left            =   4320
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Data 
      Height          =   375
      Index           =   3
      Left            =   4200
      TabIndex        =   34
      Top             =   6720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   " "
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Receita"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   10
      Left            =   120
      TabIndex        =   39
      Top             =   4440
      Width           =   1170
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Obs."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   9
      Left            =   120
      TabIndex        =   38
      Top             =   4920
      Width           =   420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Emitente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   8
      Left            =   120
      TabIndex        =   37
      Top             =   2280
      Width           =   780
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Para Escolher um Cliente Digite Seu Código, Seu Nome ou Pressione F5"
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   3360
      TabIndex        =   36
      Top             =   2880
      Width           =   4440
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Para Escolher um Tipo Monetário Digite Seu Código, Seu Nome ou Pressione F5 "
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   3360
      TabIndex        =   35
      Top             =   3840
      Width           =   4425
   End
   Begin VB.Line Line2 
      X1              =   7800
      X2              =   7800
      Y1              =   360
      Y2              =   6720
   End
   Begin VB.Label Label2 
      Caption         =   "Dados Pagamento"
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
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   6360
      Width           =   2055
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   7
      Left            =   3480
      TabIndex        =   32
      Top             =   6840
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   120
      TabIndex        =   31
      Top             =   6840
      Width           =   435
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   7800
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Monet."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   120
      TabIndex        =   30
      Top             =   3480
      Width           =   1065
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   120
      TabIndex        =   29
      Top             =   3000
      Width           =   1065
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   120
      TabIndex        =   28
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   3600
      TabIndex        =   27
      Top             =   1320
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lançamento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   120
      TabIndex        =   26
      Top             =   1320
      Width           =   1110
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Documento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   25
      Top             =   720
      Width           =   1035
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   " Controle de Contas a Receber"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   360
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   11835
   End
End
Attribute VB_Name = "Receitas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private a As Integer
Private LcLimpa As Boolean
Private LcCarregado, LcAlteradoCliente, LcAlteradoMonetario As Integer
Private Sub CmdAnterior_Click()
On Error Resume Next
GlMov = True
If MovImentacao(enAnterior, Receber) Then VinculaDados
GlMov = False
LcRegAtual = False
End Sub

Private Sub CmdAnterior_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
 Txt(0).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
  End Sub

Private Sub CmdExcluir_Click()
On Error Resume Next
GlTab = "alid015"
GlSq = "Select * from alid015 where nf='" & Txt(0).Text & "'"

If Exclui(Receber) = 1 Then
      VinculaDados
End If
LcRegAtual = False
End Sub

Private Sub CmdExcluir_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
 Txt(0).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
End Sub

Private Sub CmdFechar_Click()
On Error Resume Next
Unload frmPesquisa
Unload Me
End Sub

Private Sub CmdFechar_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
 Txt(0).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
End Sub

Private Sub CmdOrdenar_Click()
On Error Resume Next
FrmOrdena.Show , Me
End Sub

Private Sub CmdOrdenar_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
 Txt(0).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
End Sub

Private Sub CmdPesquisar_Click()
On Error Resume Next
frmPesquisa.Show , Me
LcRegAtual = False
End Sub

Private Sub CmdPesquisar_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
 Txt(0).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
End Sub

Private Sub CmdPrimeiro_Click()
On Error Resume Next
GlMov = True
If MovImentacao(enPrimeiro, Receber) Then VinculaDados
GlMov = False
LcRegAtual = False
End Sub

Private Sub CmdPrimeiro_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
 Txt(0).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
End Sub

Private Sub CmdSalvar_Click()
On Error Resume Next
Dim LcValor     As Double
Dim LcTipoMOne  As String

Call SalvaRegistro(Receber)
If GlInclusaoReceita Then
    LcTipoMOne = Txt(5).Text
    LcValor = CDbl(data(2).Text)
    Call lancacaixa("Receita", Txt(0).Text, LcTipoMOne, LcValor)
End If
VinculaDados
LcRegAtual = False
NovoReg
If LcTipoDados = 1 Then limpa
Txt(0).SetFocus
End Sub

Private Sub CmdSalvar_KeyDown(KeyCode As Integer, Shift As Integer)
 Txt(0).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
End Sub

Private Sub CmdSeguinte_Click()
On Error Resume Next
GlMov = True
If MovImentacao(enSeguinte, Receber) Then VinculaDados
GlMov = False
Txt(0).SetFocus
LcRegAtual = False
End Sub

Private Sub CmdSeguinte_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
 Txt(0).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
End Sub

Private Sub CmdUltimo_Click()
On Error Resume Next
GlMov = True
If MovImentacao(enultimo, Receber) Then VinculaDados
Txt(0).SetFocus
GlMov = False
LcRegAtual = False
End Sub



Private Sub CmdUltimo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
 Txt(0).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
End Sub

Private Sub Data_Change(Index As Integer)
If LcRegAtual Then Exit Sub

GlCampo9 = data(0).Text
GlCampo4 = data(1).Text
GlCampo20 = data(2).Text
GlCampo8 = data(3).Text
Call Alterado
End Sub

Private Sub Data_GotFocus(Index As Integer)
LcLimpa = True
End Sub

Private Sub Data_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next

 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
If KeyCode = 13 Then
   SendKeys "{TAB}"
Else
  
  End If

End Sub

Private Sub Data_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 2 Then
   If KeyAscii = 46 Then KeyAscii = 44
      If LcLimpa Then
      LcLimpa = False
      data(2).Text = ""
   End If

End If
End Sub

Private Sub Data_LostFocus(Index As Integer)
On Error Resume Next
If Index = 0 Or Index = 1 Or Index = 7 Then
   If data(Index).Text = "  /  /  " Then Exit Sub
   If Not IsDate(data(Index).Text) Then
      MsgBox "Digite Uma Data Válida.", vbInformation, "Aviso"
      data(Index).Text = "  /  /  "
      data(Index).SetFocus
      Exit Sub
   End If
End If
If Index = 2 Or Index = 3 Then
   If Len(data(Index).Text) = 0 Then Exit Sub
   If Not IsNumeric(data(Index).Text) Then
      MsgBox "Digite Um Valor Numérico.", vbInformation, "Aviso"
      data(Index).Text = ""
      data(Index).SetFocus
      Exit Sub
   End If
End If
GlCampo9 = data(0).Text
GlCampo4 = data(1).Text
GlCampo1 = data(2).Text
GlCampo8 = data(3).Text

End Sub

Private Sub Form_Activate()
On Error Resume Next
Set GlFormA = Me
If LcCarregado Then Exit Sub
Select Case LcTipoDados
   Case Is = 1
        LcCap = "   <<Modo Inclusão>>"
        data(0).Text = Format(Date, "dd/mm/yy")
        DesabilitaCtr
   Case Is = 2
       LcCap = "   <<Modo Alteração>>"
      Call AbreBanco(Receber)
      Txt(0).Locked = True
      VinculaDados
   Case Is = 3
      'DesabilitaTodos
      LcCap = "   <<Modo Consulta>>"
      MnSalvar.Enabled = False
      MnExcluir.Enabled = False
      Call AbreBanco(Receber)
      Txt(0).Locked = True
      CmdExcluir.Enabled = False
      VinculaDados
 End Select
'CriaMascara
Label1.Caption = Label1.Caption & LcCap
LcRegAtual = False
'FrmPrincipal.Visible = False
CarreGamatriz
LcCarregado = True
Txt(0).SetFocus

End Sub
Function CarreGamatriz()
Dim a As Integer, LcNome As String, LcTipo As String
GlFormAtual = Tabela.Receber
On Error Resume Next
For a = 0 To 30
   MtPesquisa(a).campo = ""
   MtPesquisa(a).Indice = ""
   MtPesquisa(a).Tipo = ""
Next

Set GlFormA = Me
For a = 0 To 5
   Select Case a
     Case Is = 0
        LcNome = "NF"
        LcCampo = "Documento"
        LcTipo = "T"
      Case Is = 1
        LcNome = "Data"
        LcCampo = "Data Lançamento"
        LcTipo = "D"
      Case Is = 2
        LcNome = "Valor"
        LcCampo = "Valor"
        LcTipo = "M"
      Case Is = 3
        LcNome = "Cliente"
        LcCampo = "Cod. Cliente"
        LcTipo = "T"
      Case Is = 4
        LcNome = "DTVENC"
        LcCampo = "Data Vencimento"
        LcTipo = "D"
      Case Is = 5
        LcNome = "DTPAGTO"
        LcCampo = "Data Pagamento"
        LcTipo = "D"
   End Select
   MtPesquisa(a).Indice = LcNome
   MtPesquisa(a).Tipo = LcTipo
   MtPesquisa(a).campo = LcCampo
 Next
 
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Call Teclas(KeyCode)
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Height = 7500
Me.Width = 10545
DataS.Text = Format(Date, "dd/mm/yyyy")
HoraS.Text = Format(Time, "hh:mm:ss")
Top = 800
Left = Screen.Width / 2 - Width / 2
LcIndice = "NF"
Call ChamaRecord("Documento")
 
End Sub
Function ChamaRecord(LcOrdem As String)
On Error Resume Next
Dim LcMat As Variant
Dim LcSql As String
Dim a As Long
LcSql = "Select * from " & Me.Name
Set RsAtual = AbreRecordset(LcSql)
RsAtual.Sort = "Codigo"
End Function

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
FechaBanco

If (LcTipoDados = 1) And (CmdSalvar.Enabled = True) Then
   GlPergunta = True
   SalvaRegistro (Receber)
End If
If (LcTipoDados = 2) And LcAlterado Then SalvaRegistro (Receber)
FechaBanco
GlStringBase = ""
GlordemAnterior = ""
FrmPrincipal.Visible = True
LcCarregado = False

End Sub




Private Sub Timer1_Timer()
On Error Resume Next
HoraS.Text = Format(Time, "hh:mm:ss")
End Sub
Private Function DesabilitaCtr()
On Error Resume Next
CmdPrimeiro.Enabled = False
CmdAnterior.Enabled = False
CmdUltimo.Enabled = False
CmdSeguinte.Enabled = False
MnMovimento.Enabled = False
MnRegistro.Enabled = False
CmdExcluir.Enabled = False
CmdPesquisar.Enabled = False
CmdOrdenar.Enabled = False
End Function
Function VinculaDados()
On Error Resume Next
If LcTipoDados = 1 Then NovoReg Else Call RegistroAtual(Receber)
If LcTipoDados = 1 Then
   GlCampo9 = Format(Date, "dd/mm/yy")
   limpa
End If
Txt(1).Text = GlCampo1
Txt(0).Text = GlCampo0
If GlCampo20 = "" Then
   data(2).Text = 0
Else
   data(2).Text = GlCampo20
End If
Txt(2).Text = GlCampo2
If GlCampo4 = "" Then
   data(1).Text = "  /  /  "
Else
   data(1).Text = Format(GlCampo4, "dd/mm/yy")
End If
Txt(5).Text = GlCampo5
Txt(7).Text = Format(GlCampo7, "dd/mm/yy")
If GlCampo8 = "" Then
   data(3).Text = 0
Else
   data(3).Text = GlCampo8
End If
If GlCampo9 = "" Then
   data(0).Text = "  /  /  "
Else
   data(0).Text = Format(GlCampo9, "dd/mm/yy")
End If
Txt(10).Text = GlCampo10
Txt(11).Text = GlCampo11
Txt(12).Text = GlCampo12

BuscaCliente (1)
BuscaTipo (1)
Txt(0).SetFocus
CmdSalvar.Enabled = False
MnSalvar.Enabled = False
If LcTipoDados = 1 Then DesabilitaCtr
Exit Function
ErroVinculo:

Resume Next
End Function

Private Sub txt_Change(Index As Integer)
On Error Resume Next
Call Alterado
If Index = 3 Then LcAlteradoCliente = True
If Index = 6 Then LcAlteradoMonetario = True
If Index = 2 Then
   GlCampo2 = Txt(2).Text
End If

End Sub


Private Sub txt_GotFocus(Index As Integer)
On Error Resume Next
If Index = 3 Then LcAlteradoCliente = False
If Index = 6 Then LcAlteradoMonetario = False
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then
   If Index <> 10 Then SendKeys "{TAB}"
Else
  If KeyCode = 116 Then
  If Index = 11 Or Index = 12 Then
      BuscaCodigoDespesa
      Exit Sub
   End If
   If Index = 2 Or Index = 3 Then
       
      GlEscolhe = 1  'Exibe Clientes
      If Len(Trim(Txt(3).Text)) > 0 Then
            'FrmBuscaCliente.txt.Text = txt(3).Text
            GlCriterioSql = " where RAZAOSOC like '" & UCase(Txt(3).Text) & "*'  order by RAZAOSOC"
       Else
            GlCriterioSql = ""
         End If
      FrmBuscaCliente.Show , Me
      Exit Sub
   Else
      If Index = 5 Or Index = 6 Then 'Exibe Produtos
         GlEscolhe = 2
         If Len(Trim(Txt(6).Text)) > 0 Then
            FrmPesquisaProdutos.Txt.Text = Txt(6).Text
            GlCriterioSql = "select * From alid008 where nome like '" & UCase(Txt(2).Text) & "*'  order by nome"
         Else
            GlCriterioSql = ""
         End If
         Teclas (KeyCode)
      End If
    End If
Else
  Teclas (KeyCode)
End If
End If
End Sub
Function limpa()
Dim a As Long
On Error Resume Next
For a = 0 To 36
  Txt(a).Text = ""
Next
data(1).Text = "  /  /  "
data(2).Text = "0"
data(3).Text = "0"
Txt(0).SetFocus
CmdSalvar.Enabled = False
End Function
Function BuscaTipo(LcTipo As Integer)
On Error GoTo erroBustaTipo
Dim RsTipo As Recordset
Dim LcDigitado, LcCodigo As String
AbreBase
Set RsTipo = Dbbase.OpenRecordset("select * from alid008", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Select Case LcTipo
    Case Is = 1 '===Chamado pelo Vincula Dados
         LcCriterioCli = "TPMONET='" & Txt(5).Text & "'"
         RsTipo.FindFirst LcCriterioCli
         If Not RsTipo.NoMatch Then
            Txt(6).Text = RsTipo!XTPMONET
            LcDesCidade = RsTipo!XTPMONET
            SendKeys "{TAB}"
         Else
            Txt(6).Text = ""
         End If
    Case Is = 2 '===Chamado Pelo Cliente
        LcValorDigitado = Txt(6).Text
        If Len(Txt(6).Text) = 0 Then Exit Function
        lcchave = Right("00" & Txt(6).Text, 2)
        LcCriterioCli = "TPMONET='" & lcchave & "'"
        RsTipo.FindFirst LcCriterioCli
        If Not RsTipo.NoMatch Then
            Txt(6).Text = RsTipo!XTPMONET
            Txt(5).Text = RsTipo!TPMONET
            LcDesCidade = RsTipo!XTPMONET
            'SendKeys "{TAB}"
        Else
            Txt(6).Text = LcValorDigitado
            If LcAlteradoMonetario Then
               ExibeMonetario.Show , Me
               LcAlteradoMonetario = False
            End If
            'Data(1).SetFocus
        End If
  
End Select

LcTipo = 0
RsTipo.Close
Set RsTipo = Nothing
Exit Function

erroBustaTipo:
If err = 3420 Then
   AbreBanco (LcTabl)
Else
   If err = 3021 Then
      Resume Next
   Else
      MsgBox err.Description & " " & err
   End If
   'Resume 0
End If

End Function
Function BuscaCliente(LcTipo As Integer)
On Error GoTo errBuscaFor
Dim RsCliente As Recordset
Dim LcValorDigitado
Dim LcCodigo As String
AbreBase
Set RsCliente = Dbbase.OpenRecordset("select * from alid001", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Select Case LcTipo
    Case Is = 1 '===Chamado pelo Vincula Dados
         LcCriterioCli = "CODIGO='" & Txt(2).Text & "'"
         RsCliente.FindFirst LcCriterioCli
         If Not RsCliente.NoMatch Then
            Txt(3).Text = RsCliente!razaosoc
            LcDesCidade = RsCliente!razaosoc
            SendKeys "{TAB}"
         Else
            Txt(3).Text = ""
         End If
    Case Is = 2 '===Chamado Pelo Cliente
        LcValorDigitado = Txt(3).Text
        If Len(Txt(3).Text) = 0 Then Exit Function
        
        lcchave = Right("00000" & Txt(3).Text, 5)
        LcCriterioCli = "CODIGO='" & lcchave & "'"
        RsCliente.FindFirst LcCriterioCli
        If Not RsCliente.NoMatch Then
            Txt(3).Text = RsCliente!razaosoc
            Txt(2).Text = RsCliente!Codigo
            LcDesCidade = RsCliente!razaosoc
            'SendKeys "{TAB}"
        Else
            Txt(3).Text = LcValorDigitado
            FrmBuscaCliente.Txt.Text = Txt(3).Text
            GlCriterioSql = " where RAZAOSOC like '" & UCase(Txt(3).Text) & "*'  order by RAZAOSOC"
            If LcAlteradoCliente Then
               FrmBuscaCliente.Show , Me
               LcAlteradoCliente = False
            End If
            'Data(1).SetFocus
        End If
  
End Select

RsCliente.Close
Set RsCliente = Nothing
Exit Function

errBuscaFor:
If err = 3420 Then
   AbreBanco (LcTabl)
Else
   If err = 3021 Then
      Resume Next
   Else
      MsgBox err.Description & " " & err
   End If
   'Resume 0
End If



End Function
Function BuscaCodigoDespesa()
Dim RsDesp As Recordset
AbreBase
If Len(Txt(11).Text) > 0 Then
   Set RsDesp = Dbbase.OpenRecordset("select * from alid007 where RD='R' and COD='" & Right("00" & Txt(11).Text, 2) & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)
   If Not RsDesp.EOF Then
      Txt(12).Text = RsDesp!Nome & ""
      GoTo ExitBusca
   Else
      Me.Tag = "D"
      exibeDespRec.Show , Me
      GoTo ExitBusca
   End If
Else
   Set RsDesp = Dbbase.OpenRecordset("select * from alid007 where RD='R' and nome='" & Txt(12).Text & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)
   If Not RsDesp.EOF Then
      Txt(11).Text = RsDesp!cod & ""
      GoTo ExitBusca
   Else
      Me.Tag = "D"
      exibeDespRec.Show , Me
      GoTo ExitBusca
   End If
End If
exibeDespRec.Show , Me

ExitBusca:
RsDesp.Close
Dbbase.Close
Exit Function
End Function
Private Sub Txt_LostFocus(Index As Integer)
If Index = 7 Then
   If Len(Txt(7).Text) = 0 Then Exit Sub
   If Not IsDate(Txt(7).Text) Then
       MsgBox "Digite uma Data Válida...", 64, "Aviso"
       Txt(7).Text = ""
       Txt(7).SetFocus
       Exit Sub
   End If
End If
 If Index = 11 Or Index = 12 Then
      BuscaCodigoDespesa
      Exit Sub
   End If
If Index = 3 Then BuscaCliente (2)
If Index = 6 Then BuscaTipo (2)
End Sub
