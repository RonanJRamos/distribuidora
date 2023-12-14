VERSION 5.00
Begin VB.Form frmPesquisaProduto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pesquisa"
   ClientHeight    =   2340
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   5445
   Icon            =   "frmPesquisaProduto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1382.549
   ScaleMode       =   0  'User
   ScaleWidth      =   5112.56
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdAnterior 
      Caption         =   "&Anterior F3"
      Enabled         =   0   'False
      Height          =   390
      Left            =   1860
      TabIndex        =   11
      Top             =   1860
      Width           =   1140
   End
   Begin VB.CommandButton CmdProximo 
      Caption         =   "&Próximo F4"
      Enabled         =   0   'False
      Height          =   390
      Left            =   3000
      TabIndex        =   10
      Top             =   1860
      Width           =   1140
   End
   Begin VB.Frame Condicao 
      Caption         =   "Condição Campo"
      Height          =   1575
      Left            =   3600
      TabIndex        =   6
      Top             =   120
      Width           =   1815
      Begin VB.OptionButton Option3 
         Caption         =   "Exatamente"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Qualquer Parte"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Inicio "
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.TextBox TxtChave 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   3375
   End
   Begin VB.ComboBox CmdPesquisa 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   360
      Width           =   2655
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK F2"
      Enabled         =   0   'False
      Height          =   390
      Left            =   720
      TabIndex        =   0
      Top             =   1860
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar F10"
      Height          =   390
      Left            =   4140
      TabIndex        =   1
      Top             =   1860
      Width           =   1260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pesquisar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Campo Para Pesquisa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1860
   End
End
Attribute VB_Name = "frmPesquisaProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LcTipoPesq As Integer, LcPosAtual As Integer
Dim LCEntrouPesquisa, a As Integer
Public LoginSucceeded As Boolean
Private RsPesquisa As ADODB.Recordset
Private CodigoAchado As String

Private Sub CmdAnterior_Click()
pesquisa 3
End Sub

Private Sub CmdAnterior_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 114 Then SendKeys "%+{A}"
If KeyCode = 115 Then SendKeys "%+{P}"
If KeyCode = 121 Then SendKeys "%+{C}"
End Sub

Private Sub cmdCancel_Click()
   On Error Resume Next
    LoginSucceeded = False
       
    Unload Me
End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 114 Then SendKeys "%+{A}"
If KeyCode = 115 Then SendKeys "%+{P}"
If KeyCode = 121 Then SendKeys "%+{C}"
End Sub

Private Sub CmdOk_Click()
pesquisa 1
End Sub

Private Sub CmdOk_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 114 Then SendKeys "%+{A}"
If KeyCode = 115 Then SendKeys "%+{P}"
If KeyCode = 121 Then SendKeys "%+{C}"
End Sub

Private Sub CmdPesquisa_Click()
On Error Resume Next
HabilitaCondicao
End Sub
Function limpa()
On Error Resume Next
cmdOK.Enabled = False
CmdProximo.Enabled = False
CmdAnterior.Enabled = False
TxtChave = ""
End Function
Private Sub CmdPesquisa_GotFocus()
limpa

End Sub

Private Sub CmdPesquisa_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then
   TxtChave.SetFocus
End If
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 114 Then SendKeys "%+{A}"
If KeyCode = 115 Then SendKeys "%+{P}"
If KeyCode = 121 Then SendKeys "%+{C}"
End Sub

Private Sub CmdProximo_Click()
pesquisa 2
End Sub



Private Sub CmdProximo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 114 Then SendKeys "%+{A}"
If KeyCode = 115 Then SendKeys "%+{P}"
If KeyCode = 121 Then SendKeys "%+{C}"
End Sub

Private Sub Form_Activate()
On Error Resume Next
If Not LCEntrouPesquisa Then
   TxtChave.SetFocus
   LCEntrouPesquisa = True
End If
HabilitaCondicao

End Sub

Private Sub Form_Load()
'Monta Combo
On Error Resume Next
CmdPesquisa.SetFocus

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim LcIndiceAntigo As String
On Error Resume Next
GLPesquisa = False
Sincroniza
LCEntrouPesquisa = False
LcRegAtual = False
GlFormA.SetFocus
End Sub

Private Sub Option1_Click()
'limpa
End Sub

Private Sub Option1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 114 Then SendKeys "%+{A}"
If KeyCode = 115 Then SendKeys "%+{P}"
If KeyCode = 121 Then SendKeys "%+{C}"
End Sub

Private Sub Option2_Click()
'limpa
End Sub

Private Sub Option2_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 114 Then SendKeys "%+{A}"
If KeyCode = 115 Then SendKeys "%+{P}"
If KeyCode = 121 Then SendKeys "%+{C}"
End Sub

Private Sub Option3_Click()
'limpa
End Sub

Private Sub Option3_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 114 Then SendKeys "%+{A}"
If KeyCode = 115 Then SendKeys "%+{P}"
If KeyCode = 121 Then SendKeys "%+{C}"
End Sub

Private Sub TxtChave_Change()
On Error Resume Next
If Len(Trim(TxtChave)) = 0 Then
   cmdOK.Enabled = False
Else
   cmdOK.Enabled = True
End If
End Sub

Private Sub TxtChave_GotFocus()
'limpa
SendKeys "{home}"
SendKeys "+{end}"
End Sub

Private Sub TxtChave_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then
   cmdOK.SetFocus
End If
If KeyCode = 113 Then SendKeys "%+{O}"
If KeyCode = 114 Then SendKeys "%+{A}"
If KeyCode = 115 Then SendKeys "%+{P}"
If KeyCode = 121 Then SendKeys "%+{C}"
End Sub

Private Sub TxtChave_LostFocus()
On Error Resume Next
TxtChave.Text = UCase(TxtChave.Text)
End Sub
Function pesquisa(LcTipoPesquisa As Integer)
On Error GoTo erroPesquisa
Dim LcSeparacao As String
Dim LcWhere     As String
Dim mark        As Variant
Dim LcAchou     As Boolean
Dim LcNome As String
Dim C As Control
Dim LcMark As String
Dim LcType As Integer

'MsgBox DEscricaoErro
Select Case LcTipoPesquisa
    Case Is = 1 '===> Solicita a Primeira Pesquisa
       
        LcComentario = "-Form:Pesquisa Function:Pesquisa - Pesquisando primeiro registro. Where=" & LcWhere
        'LcMark = RsPesquisa.Bookmark
        LcSeparacao = ""
        LcComentario = "-Form:Pesquisa Function:Pesquisa - Selecionando o Tipo do Campo"
        LcNome = CmdPesquisa.Text
        If RsPesquisa Is Nothing Then
            LcSql = "Select * from produtos order by nome "
            Set RsPesquisa = AbreRecordset(LcSql, True)
        End If
       
        LcType = RsPesquisa.Fields(LcNome).Type
        Select Case LcType
          Case Is = adDBDate
               LcComentario = "-Form:Pesquisa Function:Pesquisa - Criando where por data"
               If IsDate(TxtChave.Text) Then
                  LcWhere = LCase(C.Text) & "=#" & Format(TxtChave.Text, "mm/dd/yy") & "#"
               Else
                  MsgBox "A Data Digitada é Inválida.", 64, "Valor Não Aceito"
                  Exit Function
               End If
          Case Is = dbBoolean
               LcComentario = "-Form:Pesquisa Function:Pesquisa - Criando where por Bolean"
               If UCase(TxtChave.Text) = "SIM" Or UCase(TxtChave.Text) = "TRUE" Or TxtChave.Text = "-1" Or UCase(TxtChave.Text) = "VERDADEIRO" Or UCase(TxtChave.Text) = "VERDADE" Then
                  LcWhere = LCase(CmdPesquisa.Text) & "=True"
               Else
                  LcWhere = LCase(CmdPesquisa.Text) & "=False"
               End If
          Case adInteger
               LcComentario = "-Form:Pesquisa Function:Pesquisa - Criando where por Inteiro"
               If IsNumeric(TxtChave.Text) Then
                  LcWhere = LCase(CmdPesquisa.Text) & "=" & TxtChave.Text
               Else
                 MsgBox "O Valor Digitado não é um Valor Numérico.", 64, "Valor Não Aceito"
                 Exit Function
               End If
            Case 19
               LcComentario = "-Form:Pesquisa Function:Pesquisa - Criando where por Inteiro"
               If IsNumeric(TxtChave.Text) Then
                  LcWhere = LCase(CmdPesquisa.Text) & "=" & TxtChave.Text
               Else
                 MsgBox "O Valor Digitado não é um Valor Numérico.", 64, "Valor Não Aceito"
                 Exit Function
               End If
             
          Case Is = adNumeric
              LcComentario = "-Form:Pesquisa Function:Pesquisa - Criando where por Numeric"
              If IsNumeric(TxtChave.Text) Then
                  LcWhere = LCase(CmdPesquisa.Text) & "=" & TxtChave.Text
               Else
                 MsgBox "O Valor Digitado não é um Valor Numérico.", 64, "Valor Não Aceito"
                 Exit Function
               End If
          Case Else
               LcComentario = "-Form:Pesquisa Function:Pesquisa - Criando where por Outros tipos (String)"
               If Option1.Value Then
                  LcWhere = LCase(CmdPesquisa.Text) & " like '" & TxtChave.Text & "%'"
               Else
                  LcWhere = LCase(CmdPesquisa.Text) & " like '%" & TxtChave.Text & "%'"
               End If
        End Select
        LcComentario = "-Form:Pesquisa Function:Pesquisa - Selecionando o sentido da pesquisa  primeira,Seguinte ou anterior "
        LcSql = "Select * from produtos "
        If Len(LcWhere) > 0 Then
           LcSql = LcSql & " where " & LcWhere
        End If
        LcSql = LcSql & " order by nome "
        Set RsPesquisa = AbreRecordset(LcSql, True)
        RsPesquisa.MoveFirst
        RsPesquisa.Find LcWhere, 0, adSearchForward
        If Not RsPesquisa.EOF Then
           CodigoAchado = RsPesquisa!Codigo
        End If
        LcAchou = Not RsPesquisa.EOF
    Case Is = 2 '====> A Pesquisa Será Feita para o Proximo Registro
        'LcComentario = "-Form:Pesquisa Function:Pesquisa - Pesquisando Proximo registro. Where=" & LcWhere
       ' 'LcMark = RsPesquisa.Bookmark
       ' 'RsPesquisa.Find LcWhere, 1, adSearchForward '"Codigo=" & FrmProduto.Txt(0).Text, 1, adSearchForward
        'RsPesquisa.Find LcWhere, 1, adSearchForward
        'If Not RsPesquisa.EOF Then
        '    CodigoAchado = RsPesquisa!Codigo
        'End If
        'Do Until RsPesquisa.EOF '
        '   If RsPesquisa!Codigo = CodigoAchado Then
         '     RsPesquisa.Find LcWhere, 1, adSearchForward
         '  Else
         '     CodigoAchado = RsPesquisa!Codigo
         '     Exit Do
         '  End If
        'Loop
        If Not RsPesquisa.EOF Then
           RsPesquisa.MoveNext
        Else
           MsgBox ("Este é o ultimo Registro!")
        End If
        LcAchou = Not RsPesquisa.EOF
        'If Not RsPesquisa.EOF Then MsgBox RsPesquisa!Codigo
    Case Is = 3 '===> A Pesquisa Será para o Registro Anterior
        'LcComentario = "-Form:Pesquisa Function:Pesquisa - Pesquisando registro.Anterior Where=" & LcWhere
        ''LcMark = RsPesquisa.Bookmark
        'RsPesquisa.Find "Codigo=" & FrmProduto.Txt(0).Text, 1, adSearchForward
        'RsPesquisa.Find LcWhere, 1, adSearchBackward
       '' CodigoAchado = RsPesquisa!Codigo
        'Do Until RsPesquisa.BOF
         '  If RsPesquisa!Codigo = CodigoAchado Then
         '     RsPesquisa.Find LcWhere, 1, adSearchForward
         '  Else
         '    CodigoAchado = RsPesquisa!Codigo
          '    Exit Do
          ' End If
        'Loop
        If Not RsPesquisa.BOF Then
           RsPesquisa.MovePrevious
        Else
           MsgBox ("Este é o Primeiro Registro!")
        End If
        LcAchou = Not RsPesquisa.BOF
End Select
LcComentario = "-Form:Pesquisa Function:Pesquisa - Verificando se foi bem sucedida."
If LcAchou Then
   LcComentario = "-Form:Pesquisa Function:Pesquisa - Apresentado os dados."
   GlFormA.VinculaDados RsPesquisa!Codigo
   If LcTipoPesquisa = 1 Then
     CmdAnterior.Enabled = True
     CmdProximo.Enabled = True

   End If
Else
  ' RsPesquisa.Bookmark = LcMark
   MsgBox "Registro Não Encontrado.", 64, "Aviso"
End If

Exit Function
erroPesquisa:
MsgBox err.Description & err.Number
'Resume 0
'logErro err.Number, err.Description, LcComentario
MsgBox "Ocorreu um erro efetuando a pesquisa." & Chr(13) & "verifique os criterios de pesquisa e tente novamente.", 64, "Nº:" & err.Number & " Des:" & err.Description
Exit Function
End Function

Function CriaLista(lcform As Form, Rs As ADODB.Recordset, LcCampoAt As String)

'On Error Resume Next
Dim C As Control
Dim LcForma As Form
Dim LcNome As String

'Ordem.Clear
Set LcForma = lcform
Set RsPesquisa = Rs
For Each C In lcform.Controls()
    LcNome = UCase(C.Name)
    If Left(UCase(LcNome), 7) <> "COMMAND" Then
        If Len(C.Tag) > 0 Then
            LcNome = C.Tag
            CmdPesquisa.AddItem LcNome
        End If
    End If
Next

CmdPesquisa.Text = LcCampoAt
'TxtChave.SetFocus
'MsgBox RsPesquisa!Nome

End Function
Function HabilitaCondicao()
On Error Resume Next
Dim LcNome As String
Dim LcType As Integer

LcNome = CmdPesquisa.Text
LcType = RsPesquisa.Fields(LcNome).Type

If LcType = 200 Or LcType = adChar Or LcType = adVarChar Then
   Condicao.Enabled = True
Else
   Condicao.Enabled = False
End If
End Function
