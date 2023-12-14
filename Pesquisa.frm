VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Pesquisa 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pesquisa Registro"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11445
   Icon            =   "Pesquisa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   11445
   StartUpPosition =   3  'Windows Default
   Begin MSMask.MaskEdBox valor 
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   1545
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      _Version        =   393216
      PromptChar      =   " "
   End
   Begin VB.ListBox lista 
      Height          =   1620
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   300
      Width           =   3495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Critérios para Pesquisa"
      Height          =   975
      Left            =   3720
      TabIndex        =   1
      Top             =   240
      Width           =   3975
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Pesquisar em Qualquer Parte do Campo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   3735
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Pesquisar no Inicio do Campo"
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
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   3015
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5880
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Pesquisa.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Pesquisa.frx":0894
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Pesquisa.frx":0FA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Pesquisa.frx":13F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Pesquisa.frx":1506
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar botoes 
      Height          =   1560
      Left            =   7800
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   2752
      ButtonWidth     =   3096
      ButtonHeight    =   1376
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "&Pesquisar F2"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Pesquisar Pro&ximo F3"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Pesquisar &Anterior F4"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Cancelar F12"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor a Pesquisar"
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
      Height          =   240
      Left            =   3720
      TabIndex        =   7
      Top             =   1320
      Width           =   1875
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Campo a Pesquisar"
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
      Height          =   240
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "Pesquisa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public LcForma As Form
Private RsPesquisa As ADODB.Recordset
Function CriaLista(lcform As Form, Rs As ADODB.Recordset)
On Error Resume Next
Dim C As Control
Ordem.Clear
Set LcForma = lcform
Set RsPesquisa = Rs
For Each C In lcform.Controls()
    LcNome = UCase(C.Name)
    If LcNome <> "TITULO" And LcNome <> "BOTOES1" And LcNome <> "BARSTATUS" And LcNome <> "LINE" And LcNome <> "LABEL" And LcNome <> "TAB" And LcNome <> "BOTOES" And LcNome <> "FIGURAS" Then
        LcNome = Mid(LcNome, 1, 1) & Right(LcNome, Len(LcNome) - 1)
        Lista.AddItem LcNome
    End If
Next

'lista.Text = LcOrdem
Lista.SetFocus

End Function

Private Sub botoes_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
Select Case UCase(Button)
    Case Is = "&CANCELAR F12"
        Unload Me
    Case Is = "&PESQUISAR F2"
       Call pesquisa(1)
    Case Is = "PESQUISAR PRO&XIMO F3"
       Call pesquisa(2)
    Case Is = "PESQUISAR &ANTERIOR F4"
       Call pesquisa(3)
End Select

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then SendKeys "{TAB}"

If KeyCode = 113 Then Call botoes_ButtonClick(botoes.Buttons(1))
If KeyCode = 114 Then Call botoes_ButtonClick(botoes.Buttons(2))
If KeyCode = 115 Then Call botoes_ButtonClick(botoes.Buttons(3))
If KeyCode = 123 Then Call botoes_ButtonClick(botoes.Buttons(4))

End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Top = 300
Me.Left = 10

End Sub

Private Sub lista_Click()
On Error Resume Next
LcNome = UCase(Lista.Text)
valor.Mask = ""
valor.Text = ""
Select Case LcNome
    Case Is = "CPF"
        valor.Mask = "999.999.999-99"
    Case Is = "CNPJ"
        valor.Mask = "99.999.999/9999-99"
    Case Is = "CGC"
        valor.Mask = "99.999.999/9999-99"
    Case Is = "CEP"
        valor.Mask = "99.999-99"
    Case Else
        LcType = RsPesquisa.Fields(LcNome).Type
        If LcType = adDBDate Or LcType = adDate Then
           valor.Mask = "99/99/99"
        Else
           valor.Mask = ""
        End If
    
End Select
    
End Sub

Function pesquisa(LcTipoPesquisa As Integer)
On Error GoTo erroPesquisa
Dim LcSeparacao As String
Dim LcWhere     As String
Dim mark        As Variant
Dim LcAchou     As Boolean
LcSeparacao = ""
LcComentario = "-Form:Pesquisa Function:Pesquisa - Selecionando o Tipo do Campo"
LcNome = Lista.Text
LcType = RsPesquisa.Fields(LcNome).Type
Select Case LcType
  Case Is = adDBDate
       LcComentario = "-Form:Pesquisa Function:Pesquisa - Criando where por data"
       If IsDate(valor.Text) Then
          LcWhere = Lista.Text & "=#" & Format(valor.Text, "mm/dd/yy") & "#"
       Else
          MsgBox "A Data Digitada é Inválida.", 64, "Valor Não Aceito"
          Exit Function
       End If
  Case Is = dbBoolean
       LcComentario = "-Form:Pesquisa Function:Pesquisa - Criando where por Bolean"
       If UCase(valor.Text) = "SIM" Or UCase(valor.Text) = "TRUE" Or valor.Text = "-1" Or UCase(valor.Text) = "VERDADEIRO" Or UCase(valor.Text) = "VERDADE" Then
          LcWhere = Lista.Text & "=True"
       Else
          LcWhere = Lista.Text & "=False"
       End If
  Case adInteger
       LcComentario = "-Form:Pesquisa Function:Pesquisa - Criando where por Inteiro"
       If IsNumeric(valor.Text) Then
          LcWhere = Lista.Text & "=" & valor.Text
       Else
         MsgBox "O Valor Digitado não é um Valor Numérico.", 64, "Valor Não Aceito"
         Exit Function
       End If
  Case Is = adNumeric
      LcComentario = "-Form:Pesquisa Function:Pesquisa - Criando where por Numeric"
      If IsNumeric(valor.Text) Then
          LcWhere = Lista.Text & "=" & valor.Text
       Else
         MsgBox "O Valor Digitado não é um Valor Numérico.", 64, "Valor Não Aceito"
         Exit Function
       End If
  Case Else
       LcComentario = "-Form:Pesquisa Function:Pesquisa - Criando where por Outros tipos (String)"
       If Option1.Value Then
          LcWhere = Lista.Text & " like '" & valor.Text & "*'"
       Else
          LcWhere = Lista.Text & " like '*" & valor.Text & "*'"
       End If
End Select
LcComentario = "-Form:Pesquisa Function:Pesquisa - Selecionando o sentido da pesquisa  primeira,Seguinte ou anterior "
Select Case LcTipoPesquisa
    Case Is = 1 '===> Solicita a Primeira Pesquisa
        LcComentario = "-Form:Pesquisa Function:Pesquisa - Pesquisando primeiro registro. Where=" & LcWhere
        LcMark = RsPesquisa.Bookmark
        RsPesquisa.MoveFirst
        RsPesquisa.Find LcWhere, 0, adSearchForward
        LcAchou = Not RsPesquisa.EOF
    Case Is = 2 '====> A Pesquisa Será Feita para o Proximo Registro
        LcComentario = "-Form:Pesquisa Function:Pesquisa - Pesquisando Proximo registro. Where=" & LcWhere
        LcMark = RsPesquisa.Bookmark
        RsPesquisa.Find LcWhere, 1, adSearchForward, LcMark
        LcAchou = Not RsPesquisa.EOF
    Case Is = 3 '===> A Pesquisa Será para o Registro Anterior
        LcComentario = "-Form:Pesquisa Function:Pesquisa - Pesquisando registro.Anterior Where=" & LcWhere
        LcMark = RsPesquisa.Bookmark
        RsPesquisa.Find LcWhere, 1, adSearchBackward, LcMark
        LcAchou = Not RsPesquisa.BOF
End Select
LcComentario = "-Form:Pesquisa Function:Pesquisa - Verificando se foi bem sucedida."
If LcAchou Then
   LcComentario = "-Form:Pesquisa Function:Pesquisa - Apresentado os dados."
   VincularTabela LcForma, RsPesquisa
   'VinculaDados RsPesquisa, LcForma
   If LcTipoPesquisa = 1 Then
      botoes.Buttons(2).Enabled = True
      botoes.Buttons(3).Enabled = True

   End If
Else
   RsPesquisa.Bookmark = LcMark
   MsgBox "Registro Não Encontrado.", 64, "Aviso"
End If

Exit Function
erroPesquisa:
logErro err.Number, err.Description, LcComentario
MsgBox "Ocorreu um erro efetuando a pesquisa." & Chr(13) & "verifique os criterios de pesquisa e tente novamente.", 64, "Nº:" & err.Number & " Des:" & err.Description
Exit Function
End Function

Private Sub valor_Change()
On Error Resume Next
botoes.Buttons(1).Enabled = Len(valor.Text)
End Sub

Private Sub valor_GotFocus()
On Error Resume Next
botoes.Buttons(2).Enabled = False
botoes.Buttons(3).Enabled = False
SendKeys "{Home}"
SendKeys "+{END}"
End Sub
