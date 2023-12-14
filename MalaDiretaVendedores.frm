VERSION 5.00
Begin VB.Form MalaDiretaVendedores 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mala Direta "
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7080
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "&Sair F10"
      Height          =   495
      Left            =   3120
      TabIndex        =   6
      Top             =   2520
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar F3"
      Height          =   495
      Left            =   4920
      TabIndex        =   5
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Imprimir F2"
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   2040
      Width           =   1815
   End
   Begin VB.ComboBox chavec 
      Height          =   315
      Left            =   3120
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox Chave 
      Height          =   375
      Left            =   3120
      TabIndex        =   9
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Caption         =   "Selecionar Clientes por"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2175
      Left            =   240
      TabIndex        =   8
      Top             =   840
      Width           =   2655
      Begin VB.OptionButton estado 
         Caption         =   "Estado"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1440
         Width           =   1455
      End
      Begin VB.OptionButton cidade 
         Caption         =   "Cidade"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton nome 
         Caption         =   "Nome"
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Criterio Pesquisa"
      Height          =   375
      Left            =   3120
      TabIndex        =   10
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808000&
      Caption         =   " Imprime Mala Direta de Vendedores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "MalaDiretaVendedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private LcCampo As String
Private LcPara, a As Integer

Private Sub Chave_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 113 Then SendKeys "%{I}"
If KeyCode = 114 Then SendKeys "%{C}"
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub chavec_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 113 Then SendKeys "%{I}"
If KeyCode = 114 Then SendKeys "%{C}"
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub Cidade_Click()
LcCampo = "CIDADE"
Montacidade
Chave.Visible = False
chavec.Visible = True
End Sub

Private Sub cidade_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 113 Then SendKeys "%{I}"
If KeyCode = 114 Then SendKeys "%{C}"
If KeyCode = 121 Then SendKeys "%{F}"

End Sub

Private Sub Command1_Click()
Dim RsCliente As Recordset, RsCidade As Recordset, RsEtiqueta As Recordset
Dim LcMargem As String
Dim LcLinhaNome, LcLinhaEnd, LcLinhaBairro, LcLinhaCidade
Dim LcNomeCidade, LcEspaco As String
AbreBase
'=== Monta Criterio de Abertura da Base de dados
If LcCampo = "CIDADE" Then
   Set RsCidade = Dbbase.OpenRecordset("select  * from alid005 where NOME='" & chavec.Text & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)
   LcCriCliente = "Select * from alid200 where cidade='" & RsCidade!cod & "'"
   LcNomeCidade = RsCidade!Nome
   RsCidade.Close
   Set RsCidade = Nothing
Else
   LcCriCliente = "Select * from alid200 where " & LcCampo & " like '" & Chave.Text & "*'"
   Set RsCidade = Dbbase.OpenRecordset("alid005", dbOpenDynaset, dbSeeChanges, dbOptimistic)

End If
Set RsCliente = Dbbase.OpenRecordset(LcCriCliente, dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsEtiqueta = Dbbase.OpenRecordset("etiqueta", dbOpenDynaset, dbSeeChanges, dbOptimistic)
If RsEtiqueta.EOF Then
   MsgBox "Não Foi Configurada as Etiquetas Para mala Direta" & Chr(13) & "Entre em Utilitários/Configura Etiqueta Mala Direta", 64, "Aviso"
   GoTo SaiFuncao
End If
For sq = 1 To RsEtiqueta!LarguraHorizontal
    LcEspaco = LcEspaco & " "
Next

If RsCliente.EOF Then
   MsgBox "Não Foram Encontrados Registros com este Criterio...", 64, "Aviso"
   GoTo SaiFuncao
End If
'=== Gera Margem
For a = 1 To RsEtiqueta!MargemEsquerda
    LcMargem = LcMargem & " "
Next
'=== Abre Porta para Impressao
FnunNota = FreeFile
If Len(GlPortaMala) = 0 Then GlPortaMala = "LPT1"
Open GlPortaMala For Output Access Write As #FnunNota 'Abre Porta Nf
   
'=== Imprime Margem Superior

For a = 1 To RsEtiqueta!MargemSuperior
   Print #FnunNota, Chr(13)
Next

'=== Gera Linhas Para Imprimir
LcLinhaNome = ""
LcLinhaEnd = ""
LcLinhaBairro = ""
LcLinhaCidade = ""
Do Until RsCliente.EOF
   If LcPara Then GoTo SaiFuncao
   For c = 1 To RsEtiqueta!EtiquetaColuna
       If LcPara Then GoTo SaiFuncao
       If RsCliente.EOF Then Exit For
       For l = 1 To RsEtiqueta!EtiquetasLinha
           If LcPara Then GoTo SaiFuncao
           If RsCliente.EOF Then Exit For
           '==== Se não For a primeira então Separa pelo Espacamento Vertical
           If l > 1 Then
              For v = 1 To RsEtiqueta!DistanciaVertical
                  LcLinhaNome = LcLinhaNome & " "
                  LcLinhaEnd = LcLinhaEnd & " "
                  LcLinhaBairro = LcLinhaBairro & " "
                  LcLinhaCidade = LcLinhaCidade & " "
              Next
           End If
           If Len(LcNomeCidade) = 0 Then
              LcCri = "cod='" & RsCliente!cidade & "'"
              RsCidade.FindFirst LcCri
              If Not RsCidade.NoMatch Then
                 LcLinhaCidade = LcLinhaCidade & Left(RsCliente!Cep & "    " & RsCidade!Nome & "  " & RsCliente!estado & LcEspaco, RsEtiqueta!LarguraHorizontal)
              End If
           Else
              LcLinhaCidade = LcLinhaCidade & Left(RsCliente!Cep & "    " & LcNomeCidade & "  " & RsCliente!estado & LcEspaco, RsEtiqueta!LarguraHorizontal)
           End If
           LcLinhaNome = LcLinhaNome & Left(RsCliente!Nome & LcEspaco, RsEtiqueta!LarguraHorizontal)
           LcLinhaEnd = LcLinhaEnd & Left(RsCliente!End & LcEspaco, RsEtiqueta!LarguraHorizontal)
           LcLinhaBairro = LcLinhaBairro & Left(RsCliente!bairro & LcEspaco, RsEtiqueta!LarguraHorizontal)
           RsCliente.MoveNext
           DoEvents
       Next
       LcEsp = 0
       If RsEtiqueta!LarguraVertical > 4 Then
          LcEsp = (RsEtiqueta!LarguraVertical - 4) / 2
          LcEsp = Int(LcEsp)
       End If
       '== Centraliza a etiqueta Horizontalmente
       If LcEsp > 0 Then
          For re = 1 To LcEsp
              Print #FnunNota, Chr(13)
          Next
       End If
       Print #FnunNota, LcMargem & LcLinhaNome & Chr(13)
       Print #FnunNota, LcMargem & LcLinhaEnd & Chr(13)
       Print #FnunNota, LcMargem & LcLinhaBairro & Chr(13)
       Print #FnunNota, LcMargem & LcLinhaCidade & Chr(13)
       '==== Limpa as Variaveis para a Próxima Impressão
       
       LcLinhaNome = ""
       LcLinhaEnd = ""
       LcLinhaBairro = ""
       LcLinhaCidade = ""
 
       If LcEsp > 0 Then
          For re = 1 To LcEsp
              Print #FnunNota, Chr(13)
          Next
       End If
       LcEsp = 0
       For LcC = 1 To RsEtiqueta!DistanciaHorizontal
           Print #FnunNota, Chr(13)
       Next
       DoEvents
   Next
   DoEvents
   Print #FnunNota, Chr(12)
Loop
SaiFuncao:
Close #FnunNota
If Len(LcNomeCidade) = 0 Then
   RsCidade.Close
   Set RsCidade = Nothing
End If

RsCliente.Close
Set RsCliente = Nothing
Exit Sub
End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 113 Then SendKeys "%{I}"
If KeyCode = 114 Then SendKeys "%{C}"
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub Command2_Click()
LcPara = True
End Sub

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 113 Then SendKeys "%{I}"
If KeyCode = 114 Then SendKeys "%{C}"
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub Command3_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub Command3_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 113 Then SendKeys "%{I}"
If KeyCode = 114 Then SendKeys "%{C}"
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub estado_Click()
LcCampo = "ESTADO"
Chave.Visible = True
chavec.Visible = False
End Sub

Private Sub Fantasia_Click()
LcCampo = "FANTASIA"
Chave.Visible = True
chavec.Visible = False
End Sub

Private Sub estado_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 113 Then SendKeys "%{I}"
If KeyCode = 114 Then SendKeys "%{C}"
If KeyCode = 121 Then SendKeys "%{F}"

End Sub

Private Sub Form_Load()
LcCampo = "NOME"
Chave.Visible = True
chavec.Visible = False
LcPara = False
Montacidade
End Sub

Private Sub nome_Click()
LcCampo = "NOME"
Chave.Visible = True
chavec.Visible = False

End Sub
Function Montacidade()
Dim RsCliente As Recordset, RsCidade As Recordset, RsEtiqueta As Recordset
Dim LcMargem As String
AbreBase
'=== Monta Criterio de Abertura da Base de dados
Set RsCidade = Dbbase.OpenRecordset("alid005", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Do Until RsCidade.EOF
   chavec.AddItem RsCidade!Nome
   RsCidade.MoveNext
Loop
RsCidade.Close
Set RsCidade = Nothing
Exit Function
End Function

Private Sub Nome_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 113 Then SendKeys "%{I}"
If KeyCode = 114 Then SendKeys "%{C}"
If KeyCode = 121 Then SendKeys "%{F}"

End Sub
