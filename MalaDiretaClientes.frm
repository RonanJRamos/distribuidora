VERSION 5.00
Begin VB.Form MalaDiretaClientes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mala Direta "
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "&Arquivo F6"
      Height          =   495
      Left            =   4620
      TabIndex        =   17
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Gerar F5"
      Height          =   495
      Left            =   3240
      TabIndex        =   16
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Detalha F4"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6000
      TabIndex        =   15
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox CodigoF 
      Height          =   375
      Left            =   4920
      TabIndex        =   14
      Top             =   1320
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Codigoi 
      Height          =   375
      Left            =   3120
      TabIndex        =   13
      Top             =   1320
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Sair F10"
      Height          =   495
      Left            =   6000
      TabIndex        =   12
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar F3"
      Height          =   495
      Left            =   4620
      TabIndex        =   11
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Imprimir F2"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3240
      TabIndex        =   10
      Top             =   2400
      Width           =   1215
   End
   Begin VB.ComboBox chavec 
      Height          =   315
      Left            =   3120
      TabIndex        =   8
      Top             =   1320
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.TextBox Chave 
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   1320
      Width           =   3495
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
      Height          =   2415
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   2655
      Begin VB.OptionButton Codigo 
         Caption         =   "Código"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   2040
         Width           =   1455
      End
      Begin VB.OptionButton Fantasia 
         Caption         =   "Nome Fantasia"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton estado 
         Caption         =   "Estado"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1560
         Width           =   1455
      End
      Begin VB.OptionButton cidade 
         Caption         =   "Cidade"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   1080
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
      TabIndex        =   9
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808000&
      Caption         =   " Imprime Mala Direta de Clientes"
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
      TabIndex        =   5
      Top             =   120
      Width           =   7215
   End
End
Attribute VB_Name = "MalaDiretaClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private LcCampo As String
Private LcPara, LcDetalha As Integer
Private LcTamanho, a As Long


Private Sub Chave_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 113 Then SendKeys "%{I}"
If KeyCode = 114 Then SendKeys "%{C}"
If KeyCode = 115 Then SendKeys "%{D}"
If KeyCode = 116 Then SendKeys "%{G}"
If KeyCode = 117 Then SendKeys "%{A}"
If KeyCode = 121 Then SendKeys "%{F}"

End Sub

Private Sub chavec_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 113 Then SendKeys "%{I}"
If KeyCode = 114 Then SendKeys "%{C}"
If KeyCode = 115 Then SendKeys "%{D}"
If KeyCode = 116 Then SendKeys "%{G}"
If KeyCode = 117 Then SendKeys "%{A}"
If KeyCode = 121 Then SendKeys "%{F}"

End Sub

Private Sub Cidade_Click()
LcCampo = "CIDADE"
Montacidade
Chave.Visible = False
chavec.Visible = True
Codigoi.Visible = False
CodigoF.Visible = False
End Sub

Private Sub cidade_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 113 Then SendKeys "%{I}"
If KeyCode = 114 Then SendKeys "%{C}"
If KeyCode = 115 Then SendKeys "%{D}"
If KeyCode = 116 Then SendKeys "%{G}"
If KeyCode = 117 Then SendKeys "%{A}"
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub Codigo_Click()
LcCampo = "codigo"
Chave.Visible = False
chavec.Visible = False
Codigoi.Visible = True
CodigoF.Visible = True
End Sub

Private Sub codigo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 113 Then SendKeys "%{I}"
If KeyCode = 114 Then SendKeys "%{C}"
If KeyCode = 115 Then SendKeys "%{D}"
If KeyCode = 116 Then SendKeys "%{G}"
If KeyCode = 117 Then SendKeys "%{A}"
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub CodigoF_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 113 Then SendKeys "%{I}"
If KeyCode = 114 Then SendKeys "%{C}"
If KeyCode = 115 Then SendKeys "%{D}"
If KeyCode = 116 Then SendKeys "%{G}"
If KeyCode = 117 Then SendKeys "%{A}"
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub Codigoi_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 113 Then SendKeys "%{I}"
If KeyCode = 114 Then SendKeys "%{C}"
If KeyCode = 115 Then SendKeys "%{D}"
If KeyCode = 116 Then SendKeys "%{G}"
If KeyCode = 117 Then SendKeys "%{A}"
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub Command1_Click()
On Error GoTo errEtiqueta
Dim RsImprimeEtiqueta As Recordset, RsCidade As Recordset, RsEtiqueta As Recordset
Dim RsEtiquetaAnt As Recordset
Dim LcMargem As String
Dim LcLinhaNome, LcLinhaEnd, LcLinhaBairro, LcLinhaCidade
Dim LcNomeCidade, LcEspaco As String

Set RsEtiqueta = Dbbase.OpenRecordset("etiqueta", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsImprimeEtiqueta = Dbbase.OpenRecordset("Imprimeetiqueta", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsEtiquetaAnt = Dbbase.OpenRecordset("etiquetasAnteriores", dbOpenDynaset, dbSeeChanges, dbOptimistic)

If RsEtiqueta.EOF Then
   MsgBox "Não Foi Configurada as Etiquetas Para mala Direta" & Chr(13) & "Entre em Utilitários/Configura Etiqueta Mala Direta", 64, "Aviso"
   GoTo SaiFuncao
End If
For sq = 1 To RsEtiqueta!LarguraHorizontal
    LcEspaco = LcEspaco & " "
Next
Do Until RsImprimeEtiqueta.EOF
    RsImprimeEtiqueta.Delete
    RsImprimeEtiqueta.MoveNext
Loop


For a = 0 To LcTamanhoEtiqueta
    If MtEtiqueta(a).Imprime Then
       RsImprimeEtiqueta.AddNew
       RsImprimeEtiqueta("Nome") = MtEtiqueta(a).Nome & ""
       RsImprimeEtiqueta("End") = MtEtiqueta(a).endereco & ""
       RsImprimeEtiqueta("bairro") = MtEtiqueta(a).bairro & ""
       RsImprimeEtiqueta("cidade") = MtEtiqueta(a).cidade & ""
       RsImprimeEtiqueta("cep") = MtEtiqueta(a).Cep & ""
       RsImprimeEtiqueta("Estado") = MtEtiqueta(a).uf & ""
       RsImprimeEtiqueta.Update
    End If
Next


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
RsImprimeEtiqueta.MoveFirst
Do Until RsImprimeEtiqueta.EOF
  If LcPara Then GoTo SaiFuncao
     For c = 1 To RsEtiqueta!EtiquetaColuna
       If LcPara Then GoTo SaiFuncao
       If RsImprimeEtiqueta.EOF Then Exit For
       For l = 1 To RsEtiqueta!EtiquetasLinha
           If LcPara Then GoTo SaiFuncao
           'If RsCliente.EOF Then Exit For
           '==== Se não For a primeira então Separa pelo Espacamento Vertical
           If l > 1 Then
              For v = 1 To RsEtiqueta!DistanciaVertical
                  LcLinhaNome = LcLinhaNome & " "
                  LcLinhaEnd = LcLinhaEnd & " "
                  LcLinhaBairro = LcLinhaBairro & " "
                  LcLinhaCidade = LcLinhaCidade & " "
              Next
           End If
           
           
           If RsImprimeEtiqueta.EOF Then Exit For
           LcLinhaCidade = LcLinhaCidade & Left(RsImprimeEtiqueta!Cep & "    " & RsImprimeEtiqueta!cidade & "  " & RsImprimeEtiqueta!estado & LcEspaco, RsEtiqueta!LarguraHorizontal)
           
           LcLinhaNome = LcLinhaNome & Left(RsImprimeEtiqueta!Nome & LcEspaco, RsEtiqueta!LarguraHorizontal)
           LcLinhaEnd = LcLinhaEnd & Left(RsImprimeEtiqueta!End & LcEspaco, RsEtiqueta!LarguraHorizontal)
           LcLinhaBairro = LcLinhaBairro & Left(RsImprimeEtiqueta!bairro & LcEspaco, RsEtiqueta!LarguraHorizontal)
           RsImprimeEtiqueta.MoveNext
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
  For a = 1 To RsEtiqueta!MargemInferior
    Print #FnunNota, Chr(13)
  Next
  For a = 1 To RsEtiqueta!MargemSuperior
    Print #FnunNota, Chr(13)
  Next
Loop
SaiFuncao:
Close #FnunNota

RsImprimeEtiqueta.Close
Set RsImprimeEtiqueta = Nothing
Exit Sub
errEtiqueta:
If Err.Number = 55 Then
   MsgBox "A Impressora está sendo Utilizada no Momento...", 64, "Aviso"
   Exit Sub
End If
Resume Next
End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 113 Then SendKeys "%{I}"
If KeyCode = 114 Then SendKeys "%{C}"
If KeyCode = 115 Then SendKeys "%{D}"
If KeyCode = 116 Then SendKeys "%{G}"
If KeyCode = 117 Then SendKeys "%{A}"
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub Command2_Click()
LcPara = True
End Sub

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 113 Then SendKeys "%{I}"
If KeyCode = 114 Then SendKeys "%{C}"
If KeyCode = 115 Then SendKeys "%{D}"
If KeyCode = 116 Then SendKeys "%{G}"
If KeyCode = 117 Then SendKeys "%{A}"
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
If KeyCode = 115 Then SendKeys "%{D}"
If KeyCode = 116 Then SendKeys "%{G}"
If KeyCode = 117 Then SendKeys "%{A}"
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub Command4_Click()
LcDetalha = True
DetalhaEtiqueta.Show , Me
End Sub

Private Sub Command4_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 113 Then SendKeys "%{I}"
If KeyCode = 114 Then SendKeys "%{C}"
If KeyCode = 115 Then SendKeys "%{D}"
If KeyCode = 116 Then SendKeys "%{G}"
If KeyCode = 117 Then SendKeys "%{A}"
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub Command5_Click()
Dim RsCliente As Recordset, RsCidade As Recordset, RsEtiqueta As Recordset
LcCap = Me.Caption
AbreBase
'=== Monta Criterio de Abertura da Base de dados
If LcCampo = "CIDADE" Then
   Set RsCidade = Dbbase.OpenRecordset("select  * from alid005 where NOME='" & chavec.Text & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)
   LcCriCliente = "Select * from alid001 where cidade='" & RsCidade!cod & "'"
   LcNomeCidade = RsCidade!Nome
   RsCidade.Close
   Set RsCidade = Nothing
Else
   If LcCampo = "codigo" Then
       LcCriCliente = "Select * from alid001 where codigo >='" & Codigoi.Text & "' And codigo<='" & CodigoF.Text & "'"
   Else
       LcCriCliente = "Select * from alid001 where " & LcCampo & " like '" & Chave.Text & "*'"
   End If
   Set RsCidade = Dbbase.OpenRecordset("alid005", dbOpenDynaset, dbSeeChanges, dbOptimistic)
   
End If
Set RsCliente = Dbbase.OpenRecordset(LcCriCliente, dbOpenDynaset, dbSeeChanges, dbOptimistic)

If RsCliente.EOF Then
   MsgBox "Não Foram Encontrados Registros com este Criterio...", 64, "Aviso"
 '  GoTo SaiFuncao
End If
a = 0
qw = 1
RsCliente.MoveLast
LcQuantidade = RsCliente.RecordCount
RsCliente.MoveFirst
Do Until RsCliente.EOF
   Me.Caption = "Aguarde, Gerando Etiqueta Nº " & qw & " de " & LcQuantidade
   ReDim Preserve MtEtiqueta(a)
   LcCriterioCidade = "COD='" & RsCliente!cidade & "'"
   RsCidade.FindFirst LcCriterioCidade
   If Not RsCidade.NoMatch Then
      MtEtiqueta(a).cidade = RsCidade!Nome
   Else
      MtEtiqueta(a).cidade = " "
   End If
   MtEtiqueta(a).bairro = RsCliente!bairro & ""
   MtEtiqueta(a).Cep = RsCliente!Cep & ""
   MtEtiqueta(a).Codigo = RsCliente!Codigo & ""
   MtEtiqueta(a).endereco = RsCliente!End & ""
   MtEtiqueta(a).Nome = RsCliente!Razaosoc & ""
   MtEtiqueta(a).uf = RsCliente!estado & ""
   MtEtiqueta(a).Imprime = True
   a = a + 1
   qw = qw + 1
   If LcPara Then Exit Do
   RsCliente.MoveNext
 Loop
LcTamanhoEtiqueta = a - 1
RsCliente.Close
Set RsCliente = Nothing
Command4.Enabled = True
Command1.Enabled = True
Me.Caption = LcCap
MsgBox "Etiquetas Geradas com Sucesso...", 64, "Aviso"

End Sub

Private Sub Command5_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 113 Then SendKeys "%{I}"
If KeyCode = 114 Then SendKeys "%{C}"
If KeyCode = 115 Then SendKeys "%{D}"
If KeyCode = 116 Then SendKeys "%{G}"
If KeyCode = 117 Then SendKeys "%{A}"
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub Command6_Click()
On Error Resume Next
Dim RsAnterior As Recordset
LcCap = Me.Caption

AbreBase
a = 0
LcCriCliente = "Select * from etiquetasAnteriores"
Set RsAnterior = Dbbase.OpenRecordset(LcCriCliente, dbOpenDynaset, dbSeeChanges, dbOptimistic)
Do Until RsAnterior.EOF
   If Err.Number > 0 Then Exit Do
   ReDim Preserve MtEtiqueta(a)
   MtEtiqueta(a).bairro = RsAnterior!bairro & ""
   MtEtiqueta(a).Cep = RsAnterior!Cep & ""
   MtEtiqueta(a).Codigo = RsAnterior!Codigo & ""
   MtEtiqueta(a).endereco = RsAnterior!End & ""
   MtEtiqueta(a).Nome = RsAnterior!Nome & ""
   MtEtiqueta(a).uf = RsAnterior!estado & ""
   MtEtiqueta(a).cidade = RsAnterior!cidade & ""
   MtEtiqueta(a).Imprime = RsAnterior!Imprime
   a = a + 1
   RsAnterior.MoveNext
Loop
LcTamanhoEtiqueta = a - 1
RsAnterior.Close
Set RsAnterior = Nothing
Command4.Enabled = True
Command1.Enabled = True
Me.Caption = LcCap
MsgBox "Etiquetas Atualizadas com Sucesso...", 64, "Aviso"
End Sub

Private Sub Command6_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%{I}"
If KeyCode = 114 Then SendKeys "%{C}"
If KeyCode = 115 Then SendKeys "%{D}"
If KeyCode = 116 Then SendKeys "%{G}"
If KeyCode = 117 Then SendKeys "%{A}"
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub estado_Click()
LcCampo = "ESTADO"
Chave.Visible = True
chavec.Visible = False
Codigoi.Visible = False
CodigoF.Visible = False
End Sub

Private Sub estado_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 113 Then SendKeys "%{I}"
If KeyCode = 114 Then SendKeys "%{C}"
If KeyCode = 115 Then SendKeys "%{D}"
If KeyCode = 116 Then SendKeys "%{G}"
If KeyCode = 117 Then SendKeys "%{A}"
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub Fantasia_Click()
LcCampo = "FANTASIA"
Chave.Visible = True
chavec.Visible = False
Codigoi.Visible = False
CodigoF.Visible = False
End Sub

Private Sub Fantasia_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 113 Then SendKeys "%{I}"
If KeyCode = 114 Then SendKeys "%{C}"
If KeyCode = 115 Then SendKeys "%{D}"
If KeyCode = 116 Then SendKeys "%{G}"
If KeyCode = 117 Then SendKeys "%{A}"
If KeyCode = 121 Then SendKeys "%{F}"
End Sub

Private Sub Form_Load()
LcCampo = "RAZAOSOC"
Chave.Visible = True
chavec.Visible = False
LcPara = False
Montacidade
ReDim MtEtiqueta(0)
End Sub

Private Sub nome_Click()
LcCampo = "RAZAOSOC"
Chave.Visible = True
chavec.Visible = False
Codigoi.Visible = False
CodigoF.Visible = False
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
If KeyCode = 115 Then SendKeys "%{D}"
If KeyCode = 116 Then SendKeys "%{G}"
If KeyCode = 117 Then SendKeys "%{A}"
If KeyCode = 121 Then SendKeys "%{F}"
End Sub
