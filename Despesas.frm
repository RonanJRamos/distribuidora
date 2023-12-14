VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form Despesas 
   BackColor       =   &H00CAE1A2&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Controle de Contas a Pagar"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Codigo 
      Height          =   375
      Left            =   6240
      TabIndex        =   38
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Txt 
      Height          =   375
      Index           =   11
      Left            =   1320
      TabIndex        =   8
      Top             =   3600
      Width           =   1575
   End
   Begin VB.TextBox Txt 
      Height          =   375
      Index           =   12
      Left            =   3000
      TabIndex        =   9
      Top             =   3600
      Width           =   4575
   End
   Begin VB.TextBox Txt 
      Height          =   1455
      Index           =   10
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   4200
      Width           =   6255
   End
   Begin VB.CommandButton CmdFechar 
      BackColor       =   &H00D8C5B6&
      Caption         =   "&Fechar F10"
      Height          =   615
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3870
      Width           =   2385
   End
   Begin VB.CommandButton CmdSalvar 
      BackColor       =   &H00D8C5B6&
      Caption         =   "&Salvar F2"
      Enabled         =   0   'False
      Height          =   615
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1320
      Width           =   1185
   End
   Begin VB.CommandButton CmdExcluir 
      BackColor       =   &H00D8C5B6&
      Caption         =   "&Excluir F3"
      Height          =   615
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1320
      Width           =   1185
   End
   Begin VB.CommandButton CmdAnterior 
      BackColor       =   &H00D8C5B6&
      Caption         =   "&Anterior F7"
      Height          =   615
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2640
      Width           =   1185
   End
   Begin VB.CommandButton CmdSeguinte 
      BackColor       =   &H00D8C5B6&
      Caption         =   "Se&guinte F8"
      Height          =   615
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3255
      Width           =   1185
   End
   Begin VB.CommandButton CmdUltimo 
      BackColor       =   &H00D8C5B6&
      Caption         =   "&Ultimo F9"
      Height          =   615
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3255
      Width           =   1185
   End
   Begin VB.CommandButton CmdPrimeiro 
      BackColor       =   &H00D8C5B6&
      Caption         =   "&Primeiro F6"
      Height          =   615
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2640
      Width           =   1185
   End
   Begin VB.CommandButton CmdOrdenar 
      BackColor       =   &H00D8C5B6&
      Caption         =   "&Ordenar F12"
      Height          =   615
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1935
      Width           =   1185
   End
   Begin VB.CommandButton CmdPesquisar 
      BackColor       =   &H00D8C5B6&
      Caption         =   "Pes&quisa F11"
      Height          =   615
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1935
      Width           =   1185
   End
   Begin VB.TextBox Txt 
      Height          =   375
      Index           =   7
      Left            =   8640
      MaxLength       =   8
      TabIndex        =   11
      TabStop         =   0   'False
      Tag             =   "S/D/S/07/N/DTPAGTO"
      Top             =   4920
      Width           =   1575
   End
   Begin VB.TextBox Txt 
      Height          =   375
      Index           =   6
      Left            =   1320
      TabIndex        =   7
      Top             =   2520
      Width           =   6255
   End
   Begin MSMask.MaskEdBox Data 
      Height          =   375
      Index           =   0
      Left            =   1320
      TabIndex        =   1
      Top             =   1080
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
      Left            =   6120
      TabIndex        =   6
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
      Top             =   1560
      Width           =   6255
   End
   Begin VB.TextBox Txt 
      Enabled         =   0   'False
      Height          =   375
      Index           =   2
      Left            =   6720
      TabIndex        =   3
      Tag             =   "S/T/S/02/N/CREDOR"
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
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
      Left            =   8160
      TabIndex        =   22
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
      Left            =   8160
      TabIndex        =   21
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
      TabIndex        =   5
      Top             =   2040
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
      Top             =   1080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Data 
      Height          =   375
      Index           =   3
      Left            =   8640
      TabIndex        =   33
      Top             =   5295
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   " "
   End
   Begin VB.TextBox Txt 
      Height          =   375
      Index           =   0
      Left            =   1320
      MaxLength       =   25
      TabIndex        =   0
      Tag             =   "S/T/S/00/N/NF"
      Top             =   600
      Width           =   4575
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   2400
      Picture         =   "Despesas.frx":0000
      Top             =   3000
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   2400
      Picture         =   "Despesas.frx":0CCA
      Top             =   2040
      Width           =   480
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      X1              =   7800
      X2              =   10440
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Desp."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   37
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Obs"
      BeginProperty Font 
         Name            =   "Arial"
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
      TabIndex        =   36
      Top             =   4200
      Width           =   360
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Para Escolher um Tipo Monetário Digite Seu Código, Seu Nome ou Pressione F5 "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   2880
      TabIndex        =   35
      Top             =   3000
      Width           =   4785
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Para Escolher um Fornecedor Digite Seu Código, Seu Nome ou Pressione F5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   2880
      TabIndex        =   34
      Top             =   2040
      Width           =   4800
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      X1              =   7800
      X2              =   7800
      Y1              =   360
      Y2              =   6480
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Dados Pagamento"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   7920
      TabIndex        =   32
      Top             =   4680
      Width           =   2295
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor"
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
      Index           =   7
      Left            =   7920
      TabIndex        =   31
      Top             =   5400
      Width           =   510
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data"
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
      Index           =   6
      Left            =   7920
      TabIndex        =   30
      Top             =   5040
      Width           =   435
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      X1              =   0
      X2              =   7800
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Monet."
      BeginProperty Font 
         Name            =   "Arial"
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
      TabIndex        =   29
      Top             =   2640
      Width           =   1020
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento"
      BeginProperty Font 
         Name            =   "Arial"
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
      TabIndex        =   28
      Top             =   2160
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fornecedor"
      BeginProperty Font 
         Name            =   "Arial"
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
      TabIndex        =   27
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor"
      BeginProperty Font 
         Name            =   "Arial"
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
      TabIndex        =   26
      Top             =   1080
      Width           =   450
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lançamento"
      BeginProperty Font 
         Name            =   "Arial"
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
      TabIndex        =   25
      Top             =   1200
      Width           =   1065
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Documento"
      BeginProperty Font 
         Name            =   "Arial"
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
      TabIndex        =   24
      Top             =   720
      Width           =   990
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   " Controle de Contas a Pagar"
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
      TabIndex        =   23
      Top             =   0
      Width           =   11835
   End
End
Attribute VB_Name = "Despesas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private LcCarregado, LcAlteradoMonetario, LcAlteradoFornecedor, a As Integer
Private LcLimpa As Boolean
Private Rs As DAO.Recordset

Private Sub CmdAnterior_Click()
On Error Resume Next
GlMov = True
'If MovImentacao(enAnterior, pagar) Then VinculaDados
If Not Rs.BOF Then
  Rs.MovePrevious
  VinculaDados
End If
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

'GlTab = "Alid014"
'GlSq = "Select * from alid014 where codigo=" & Codigo.Text
AbreBase
Dbbase.Execute "Delete from alid014 where codigo=" & Codigo.Text
Set Rs = Dbbase.OpenRecordset("Select * from alid014 order by codigo")
VinculaDados

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
CarreGamatriz
Set GlFormA = Me
LcIndice = "NF"
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
'If MovImentacao(enPrimeiro, pagar) Then VinculaDados
Rs.MoveFirst
VinculaDados
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
Public Sub pesquisa(criterio As String, Tipo As Integer)
Select Case Tipo
    Case Is = 0
         Rs.FindFirst criterio
    Case Is = 1
        Rs.FindNext criterio
    Case Is = 2
        Rs.FindPrevious criterio
End Select
VinculaDados
End Sub
Private Sub CmdSalvar_Click()
'On Error Resume Next
Dim LcTipoMOne  As String
Dim LcValor     As Double
Dim LcCodigo As String
LcCodigo = Codigo.Text
Dim RsBase As DAO.Recordset
Dim Erro As String

'===> Valida
If Len(Txt(0).Text) = 0 Then
   Erro = "Informe o Numero do Documento."
End If
If Not IsDate(Data(0).Text) Then
   If Len(Erro) > 0 Then Erro = Erro & Chr(13)
   Erro = Erro & "Informe a data de lançamento."
End If
If Not IsNumeric(Data(2).Text) Then
   If Len(Erro) > 0 Then Erro = Erro & Chr(13)
   Erro = Erro & "Informe o valor do ducumento."
Else
   If CDec(Data(2).Text) = 0 Then
        If Len(Erro) > 0 Then Erro = Erro & Chr(13)
        Erro = Erro & "Informe o valor do ducumento."
   End If
End If
If Len(Txt(3).Text) = 0 Then
   If Len(Erro) > 0 Then Erro = Erro & Chr(13)
   Erro = Erro & "Informe o fornecesdor."
End If
If Not IsDate(Data(1).Text) Then
   If Len(Erro) > 0 Then Erro = Erro & Chr(13)
   Erro = Erro & "Informe a data de vencimento."
End If

If Len(Txt(6).Text) = 0 Then
   If Len(Erro) > 0 Then Erro = Erro & Chr(13)
   Erro = Erro & "Informe o tipo monetário."
End If
If Len(Erro) > 0 Then
   MsgBox Erro, 64, "Erro"
   Exit Sub
End If
AbreBase
If IsNumeric(LcCodigo) Then
   Set RsBase = Dbbase.OpenRecordset("select * from Alid014 where codigo=" & LcCodigo)
   RsBase.Edit
Else
   Set RsBase = Dbbase.OpenRecordset("Alid014")
   RsBase.AddNew
   Codigo.Text = RsBase!Codigo
End If
'Call SalvaRegistro(pagar)
If Len(Txt(5).Text) = 0 Then Txt(5).Text = "03"
RsBase!NF = Txt(0).Text
If IsNumeric(Data(2).Text) Then
   RsBase!Valor = Data(2).Text
Else
   RsBase!Valor = 0
End If
RsBase!credor = Txt(2).Text 'VerificaTipo(2, GlCampo2)
If IsDate(Data(1).Text) Then RsBase!DTVENC = Data(1).Text
RsBase!TPMONET = Txt(5).Text
If IsDate(Txt(7).Text) Then RsBase!DTPAGTO = Txt(7).Text
If IsNumeric(Data(3).Text) Then RsBase!VALPAGO = Data(3).Text Else RsBase!VALPAGO = 0

If IsDate(Data(0).Text) Then RsBase!Data = Data(0).Text
RsBase!Obs = Txt(10).Text
RsBase!codDesp = Txt(11).Text
RsBase!NomeDesp = Txt(12).Text
            
RsBase.Update



If GlInclusaoDespesa Then
    LcTipoMOne = Txt(5).Text
    LcValor = CDbl(Data(2).Text)
    Call lancacaixa("Despesas", Txt(0).Text, LcTipoMOne, LcValor)
End If
VinculaDados
LcRegAtual = False
NovoReg
If LcTipoDados = 1 Then limpa
Txt(0).SetFocus
End Sub

Private Sub CmdSalvar_KeyDown(KeyCode As Integer, Shift As Integer)
 On Error Resume Next
 Txt(0).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
End Sub

Private Sub CmdSeguinte_Click()
On Error Resume Next
GlMov = True
'If MovImentacao(enSeguinte, pagar) Then VinculaDados
If Not Rs.EOF Then
  Rs.MoveNext
  VinculaDados
End If
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
'If MovImentacao(enultimo, pagar) Then VinculaDados
Rs.MoveLast
VinculaDados
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
CmdSalvar.Enabled = True
If LcRegAtual Then Exit Sub

'GlCampo9 = Data(0).Text
'GlCampo4 = Data(1).Text
'GlCampo1 = Data(2).Text
'GlCampo8 = Data(3).Text
'Call Alterado
End Sub

Private Sub Data_GotFocus(Index As Integer)
LcLimpa = True

End Sub

Private Sub Data_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then
   SendKeys "{TAB}"
Else
   Txt(1).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
End If
End Sub

Private Sub Data_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 2 Then
   If KeyAscii = 46 Then KeyAscii = 44
   If LcLimpa Then
      LcLimpa = False
      Data(2).Text = ""
   End If
End If
End Sub

Private Sub Data_LostFocus(Index As Integer)
On Error Resume Next
If Index = 0 Or Index = 1 Or Index = 7 Then
   If Data(Index).Text = "  /  /  " Then Exit Sub
   If Not IsDate(Data(Index).Text) Then
      MsgBox "Digite Uma Data Válida.", vbInformation, "Aviso"
      Data(Index).Text = ""
      Data(Index).SetFocus
      Exit Sub
   End If
End If
If Index = 2 Or Index = 3 Then
   If Len(Data(Index).Text) = 0 Then Exit Sub
   If Not IsNumeric(Data(Index).Text) Then
      MsgBox "Digite Um Valor Numérico.", vbInformation, "Aviso"
      Data(Index).Text = ""
      Data(Index).SetFocus
      Exit Sub
   End If
End If
GlCampo9 = Data(0).Text
GlCampo4 = Data(1).Text
GlCampo1 = Data(2).Text
GlCampo8 = Data(3).Text

End Sub

Private Sub DataS_KeyDown(KeyCode As Integer, Shift As Integer)
 Txt(1).SetFocus
 If KeyCode = 116 Then
  Else
    Call Teclas(KeyCode)
  End If
End Sub

Private Sub Form_Activate()
On Error Resume Next
Set GlFormA = Me
If LcCarregado Then Exit Sub
Select Case LcTipoDados
   Case Is = 1
        LcCap = "   <<Modo Inclusão>>"
        DesabilitaCtr
        Data(0).Text = Format(Date, "dd/mm/yy")
   Case Is = 2
      LcCap = "   <<Modo Alteração>>"
      'Call AbreBanco(pagar)
      Txt(0).Locked = True
      VinculaDados
   Case Is = 3
      LcCap = "   <<Modo Consulta>>"
      MnSalvar.Enabled = False
      MnExcluir.Enabled = False
      'Call AbreBanco(pagar)
      CmdExcluir.Enabled = False
      Txt(0).Locked = True
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
GlFormAtual = Tabela.pagar
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
        LcNome = "CREDOR"
        LcCampo = "Cod. Fornecedor"
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

Private Sub Form_Load()
On Error Resume Next
'Me.Height = 7095
'Me.Width = 10545
DataS.Text = Format(Date, "dd/mm/yyyy")
HoraS.Text = Format(Time, "hh:mm:ss")
Top = 800
Left = Screen.Width / 2 - Width / 2
LcIndice = "NF"
AbreBase
Set Rs = Dbbase.OpenRecordset("Select * from alid014 order by codigo")
 If LcTipoDados <> 1 Then
    VinculaDados
  Else
    Data(0).Text = Format(Date, "dd/mm/yy")
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
'FechaBanco

'If (LcTipoDados = 1) And (CmdSalvar.Enabled = True) Then
'   GlPergunta = True
'   SalvaRegistro (pagar)
'End If
'If (LcTipoDados = 2) And LcAlterado Then SalvaRegistro (pagar)
'FechaBanco
'GlStringBase = ""
'GlordemAnterior = ""
'FrmPrincipal.Visible = True
'LcCarregado = False
'FrmPrincipal.SetFocus
End Sub




Private Sub Timer1_Timer()
On Error Resume Next
HoraS.Text = Format(Time, "hh:mm:ss")
End Sub
Private Function DesabilitaCtr()
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
'If LcTipoDados = 1 Then NovoReg Else Call RegistroAtual(pagar)
If LcTipoDados = 1 Then
   'GlCampo9 = Format(Date, "dd/mm/yy")
   limpa
End If
If Rs.EOF Or Rs.BOF Then Exit Function
If Len(Trim(Rs!NF)) = 0 Then
    Txt(0).Text = 0
 Else
    Txt(0).Text = Rs!NF & ""
 End If
 If IsNumeric(Rs!Valor) Then
    Data(2).Text = FormatNumber(Rs!Valor, 2) & ""
 Else
    Data(2).Text = 0
 End If

Txt(2).Text = Rs!credor & ""
If IsDate(Rs!DTVENC) Then
   Data(1).Text = Format(Rs!DTVENC, "dd/mm/yy")
Else
   Data(1).Text = "  /  /  "
End If

Txt(5).Text = Rs!TPMONET & ""

If IsDate(Rs!DTPAGTO) Then
   Txt(7).Text = Format(Rs!DTPAGTO, "dd/mm/yy")
Else
   Txt(7).Text = "  /  /  "
End If
If IsNumeric(Rs!VALPAGO) Then
    Data(3).Text = FormatNumber(Rs!VALPAGO, 2) & ""
 Else
    Data(3).Text = 0
 End If
 
 If IsDate(Rs!Data) Then
   Data(0).Text = Format(Rs!Data, "dd/mm/yy")
Else
   Data(0).Text = Format(Date, "dd/mm/yy")
End If
Txt(10).Text = Rs!Obs & ""
Txt(11).Text = Rs!codDesp & ""
Txt(12).Text = Rs!NomeDesp & ""

Codigo.Text = Rs!Codigo
BuscaFornecedor (1)
BuscaTipo (1)
Txt(0).SetFocus
CmdSalvar.Enabled = False
MnSalvar.Enabled = False
LcRegAtual = False
Exit Function
ErroVinculo:

Resume Next
End Function

Private Sub Txt_Change(Index As Integer)
CmdSalvar.Enabled = True
'Call Alterado
'If Index = 3 Then LcAlteradoFornecedor = True
'If Index = 6 Then LcAlteradoMonetario = True
'If Index = 2 Then
'   GlCampo2 = Txt(2).Text
'End If

End Sub


Private Sub txt_GotFocus(Index As Integer)
If Index = 3 Then LcAlteradoFornecedor = False
If Index = 6 Then LcAlteradoMonetario = False
End Sub
Function BuscaCodigoDespesa()
Dim RsDesp As Recordset
AbreBase
If Len(Txt(11).Text) > 0 Then
   Set RsDesp = Dbbase.OpenRecordset("select * from alid007 where RD='D' and COD='" & Right("00" & Txt(11).Text, 2) & "'")
   If Not RsDesp.EOF Then
      Txt(12).Text = RsDesp!Nome & ""
      GoTo ExitBusca
   Else
      Me.Tag = "D"
      exibeDespRec.Show , Me
      GoTo ExitBusca
   End If
Else
   Set RsDesp = Dbbase.OpenRecordset("select * from alid007 where RD='D' and nome='" & Txt(12).Text & "'")
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
Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then
   If Index <> 10 Then
      SendKeys "{TAB}"
   End If
Else
  If KeyCode = 116 Then
   If Index = 11 Or Index = 12 Then
      BuscaCodigoDespesa
      Exit Sub
   End If
      
   If Index = 2 Or Index = 3 Then
      GlEscolhe = 1  'Exibe Clientes
      If Len(Trim(Txt(3).Text)) > 0 Then
            FrmPesquisaFornecedor.Txt.Text = Txt(3).Text
            GlCriterioSql = "select * From alid002 where RAZAOSOC like '" & UCase(Txt(3).Text) & "*'  order by RAZAOSOC"
         Else
            GlCriterioSql = ""
         End If
      Teclas (KeyCode)
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

Data(1).Text = "  /  /  "
Data(2).Text = "0"
Data(3).Text = "0"
Data(0).Text = Format(Date, "dd/mm/yy")
Codigo.Text = ""
CmdSalvar.Enabled = False

Txt(0).SetFocus

End Function
Function BuscaFornecedor(LcTipo As Integer)
On Error GoTo errBuscaFor
Dim RsFornecedor As Recordset
Dim LcValorDigitado
Dim LcCodigo As String
AbreBase
Set RsFornecedor = Dbbase.OpenRecordset("select * from alid002")
Select Case LcTipo
    Case Is = 1 '===Chamado pelo Vincula Dados
         LcCriterioCli = "CODIGO='" & Txt(2).Text & "'"
         RsFornecedor.FindFirst LcCriterioCli
         If Not RsFornecedor.NoMatch Then
            Txt(3).Text = RsFornecedor!RazaoSoc
            LcDesCidade = RsFornecedor!RazaoSoc
            SendKeys "{TAB}"
         Else
            Txt(3).Text = ""
         End If
    Case Is = 2 '===Chamado Pelo Cliente
        LcValorDigitado = Txt(3).Text
        If Len(Txt(3).Text) = 0 Then Exit Function
        
        lcchave = Right("00000" & Txt(3).Text, 5)
        LcCriterioCli = "CODIGO='" & lcchave & "'"
        RsFornecedor.FindFirst LcCriterioCli
        If Not RsFornecedor.NoMatch Then
            Txt(3).Text = RsFornecedor!RazaoSoc
            Txt(2).Text = RsFornecedor!Codigo
            LcDesCidade = RsFornecedor!RazaoSoc
            'SendKeys "{TAB}"
        Else
            Txt(3).Text = LcValorDigitado
            FrmPesquisaFornecedores.Txt.Text = Txt(3).Text
            GlCriterioSql = "select * From alid002 where RAZAOSOC like '" & UCase(Txt(3).Text) & "*'  order by RAZAOSOC"
            If LcAlteradoFornecedor Then
               FrmPesquisaFornecedores.Show , Me
               LcAlteradoFornecedor = False
            End If
            'Data(1).SetFocus
        End If
  
End Select

RsFornecedor.Close
Set RsFornecedor = Nothing
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
Function BuscaTipo(LcTipo As Integer)
On Error GoTo erroBustaTipo
Dim RsTipo As Recordset
Dim LcDigitado, LcCodigo As String
AbreBase
Set RsTipo = Dbbase.OpenRecordset("select * from alid008")
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

Private Sub Txt_LostFocus(Index As Integer)
On Error Resume Next
 If Index = 11 Or Index = 12 Then
      BuscaCodigoDespesa
      Exit Sub
End If
If Index = 7 Then
   If Len(Txt(7).Text) = 0 Then Exit Sub
   If Not IsDate(Txt(7).Text) Then
       MsgBox "Digite uma Data Válida...", 64, "Aviso"
       Txt(7).Text = ""
       Txt(7).SetFocus
       Exit Sub
   End If
End If
If Index = 3 Then BuscaFornecedor (2)
If Index = 6 Then BuscaTipo (2)

End Sub
