VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form RelEntradaProdutos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatorio de Entrada de Produtos"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4770
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   4770
   StartUpPosition =   3  'Windows Default
   Begin MSMask.MaskEdBox datai 
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   2280
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yy"
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin VB.TextBox Copias 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3240
      TabIndex        =   11
      Text            =   "1"
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin Crystal.CrystalReport Relatorio 
      Left            =   2280
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton CmdFechar 
      Caption         =   "&Fechar"
      Height          =   615
      Left            =   2400
      TabIndex        =   15
      Top             =   3120
      Width           =   2295
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      Height          =   615
      Left            =   120
      TabIndex        =   14
      Top             =   3120
      Width           =   2175
   End
   Begin VB.TextBox Txt 
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   2640
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.TextBox Txt 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   2640
      Width           =   4575
   End
   Begin VB.Frame Frame3 
      Caption         =   "Saída"
      Height          =   975
      Left            =   3240
      TabIndex        =   21
      Top             =   120
      Width           =   1335
      Begin VB.OptionButton Impressora 
         Caption         =   "Impressora"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton Vídeo 
         Caption         =   "Vídeo"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo Comparação"
      Height          =   1695
      Left            =   1560
      TabIndex        =   20
      Top             =   120
      Width           =   1575
      Begin VB.OptionButton Opt6 
         Caption         =   "Iniciado Por"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   1455
      End
      Begin VB.OptionButton Opt5 
         Caption         =   "Menor Igual"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   1215
      End
      Begin VB.OptionButton Opt4 
         Caption         =   "Maior Igual"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   1215
      End
      Begin VB.OptionButton Opt3 
         Caption         =   "Maior"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   735
      End
      Begin VB.OptionButton Opt2 
         Caption         =   "Menor"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton Opt1 
         Caption         =   "Igual"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filtrar"
      Height          =   1695
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   1335
      Begin VB.OptionButton Bairro 
         Caption         =   "Período"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   975
      End
      Begin VB.OptionButton Endereço 
         Caption         =   "Código"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton Nome 
         Caption         =   "Produto"
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin MSMask.MaskEdBox Dataf 
      Height          =   375
      Left            =   3000
      TabIndex        =   23
      Top             =   2280
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yy"
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin VB.Label titulo1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº Copias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   3240
      TabIndex        =   18
      Top             =   1080
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label Titulo2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Titulo2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2160
      TabIndex        =   17
      Top             =   2280
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label titulo1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   1920
      Width           =   705
   End
End
Attribute VB_Name = "RelEntradaProdutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private LcEscolha, LcDetalhe, LcCampo, LcExpressao, LcCap1 As String

Private Sub Atraso_Click()
Call Exibicao(2)
titulo1(0).Caption = "Dia Limite"
Titulo2.Visible = False
Txt(1).Visible = False
LcCampo = "Descricao"
LcDetalhe = "'"
End Sub

Private Sub Bairro_Click()
Opt1.Visible = True
Opt1.Caption = "Entre"
Opt2.Visible = False
Opt3.Visible = False
Opt4.Visible = False
Opt5.Visible = False

titulo1(0).Caption = "Periodo"
Titulo2.Caption = "Até"
Titulo2.Visible = True
Txt(0).Visible = False
Txt(1).Visible = False
Datai.Visible = True
Dataf.Visible = True
Datai.SetFocus
LcCampo = "datacompra"
LcDetalhe = ""
End Sub

Private Sub Cidade_Click()
Call Exibicao(1)
titulo1(0).Caption = "Cidade"
Titulo2.Visible = False
Txt(1).Visible = False
LcCampo = "Cidade"
LcDetalhe = "'"
End Sub

Private Sub CmdFechar_Click()
On Error Resume Next
Unload Me
End Sub


Function Exibicao(LcTipo As Integer)
'LcTipo igual a 1 Será do Tipo String e doi numero e data
On Error Resume Next
Opt1.Visible = True
Opt2.Visible = True
Opt3.Visible = True
Opt4.Visible = True
Opt5.Visible = True
Opt6.Visible = True

If LcTipo = 1 Then

   Opt1.Caption = "Igual"
   Opt2.Caption = "Iniciado Por"
   Opt3.Caption = "Que Tenha"
   Opt4.Visible = False
   Opt5.Visible = False
   Opt6.Visible = False
Else
   Opt1.Caption = "Igual"
   Opt2.Caption = "Menor"
   Opt3.Caption = "Maior"
   Opt4.Caption = "Menor Igual"
   Opt5.Caption = "Maior Igual"
   Opt6.Caption = "Entre"
End If
End Function

Private Sub CmdOk_Click()
'Efetua a Selecao Campo
Dim LcFormula As String
'On Error Resume Next
Dim RsEmpresa As Recordset
Dim a, item, LcResposta As Long
Dim LcCriterio, LcEmpresa, LcEndereco, LcFone As String


Set Dbbase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
Set RsEmpresa = Dbbase.OpenRecordset("Empresa", dbOpenDynaset)

LcEmpresa = RsEmpresa!Razao
LcEndereco = RTrim(RsEmpresa!Endereco) & " Bairro: " & RsEmpresa!Bairro & "  Cidade: " & RsEmpresa!Cidade
LcFone = "Fone: " & RsEmpresa!Fone
If Not IsNull(RsEmpresa!Fax) Then
   LcFone = LcFone & " Fax: " & RsEmpresa!Fax
End If
If err <> 0 Then
 LcEmpresa = ""
 LcFone = ""
 LcEndereco = ""
End If
RsEmpresa.Close
If LcCampo = "datacompra" Then
         strData = CDate(Format(Datai.Text, "dd/mm/yyyy"))
         LcAno = Year(strData)
         LcMes = Month(strData)
         LcDia = Day(strData)
         LcDataInicio = LcAno & "," & LcMes & "," & LcDia
         LcChav1 = " date(" & LcDataInicio & ")"
         
         strData = CDate(Format(Dataf.Text, "dd/mm/yyyy"))
         LcAno = Year(strData)
         LcMes = Month(strData)
         LcDia = Day(strData)
         LcDataInicio = LcAno & "," & LcMes & "," & LcDia
         LcChav2 = " date(" & LcDataInicio & ")"
         LcFormula = "{EntradaProduto.datacompra} >=" & LcChav1 & " And {EntradaProduto.datacompra} <=" & LcChav2
        lctitulo = "Entrada de Produtos Período " & Format(Datai.Text, "dd/mm/yyyy") & " Até " & Format(Dataf.Text, "dd/mm/yyyy")
Else
   lctitulo = "Entrada de Produtos"
  Select Case LcExpressao
       Case Is = "BETWEEN"
          LcFormula = "{EntradaProduto." & LcCampo & "} " & LcExpressao & " " & LcDetalhe & Txt(0).Text & LcDetalhe & " And " & LcDetalhe & Txt(1).Text & LcDetalhe
       Case Is = "Like"
          If LcEscolha = "Iniciado Por" Then
             LcFormula = "{EntradaProduto." & LcCampo & "} " & LcExpressao & " " & LcDetalhe & UCase(Txt(0).Text) & "*" & LcDetalhe _
             & " or {EntradaProduto." & LcCampo & "} " & LcExpressao & " " & LcDetalhe & LCase(Txt(0).Text) & "*" & LcDetalhe
          Else
             LcFormula = "{EntradaProduto." & LcCampo & "} " & LcExpressao & " " & LcDetalhe & "*" & UCase(Txt(0).Text) & "*" & LcDetalhe _
             & " or {EntradaProduto." & LcCampo & "} " & LcExpressao & " " & LcDetalhe & "*" & LCase(Txt(0).Text) & "*" & LcDetalhe
          End If
       Case Else
           LcFormula = "{EntradaProduto." & LcCampo & "} " & LcExpressao & " " & LcDetalhe & UCase(Txt(0).Text) & LcDetalhe _
            & " Or {EntradaProduto." & LcCampo & "} " & LcExpressao & " " & LcDetalhe & LCase(Txt(0).Text) & LcDetalhe
  End Select
End If
' MsgBox LcFormula
Relatorio.DataFiles(0) = GLBase
Relatorio.ReportFileName = App.Path & "\EntradaProduto.rpt"
Relatorio.SelectionFormula = LcFormula

Relatorio.SortFields(0) = "+{EntradaProduto." & LcCampo & "}"
Relatorio.CopiesToPrinter = Val(copias.Text)


Relatorio.WindowTop = 50
Relatorio.WindowWidth = 700
Relatorio.WindowLeft = 50
Relatorio.WindowHeight = 500
Relatorio.WindowTitle = "Entrada de Produtos "

Relatorio.Formulas(0) = "titulo='" & lctitulo & "'"
Relatorio.Formulas(1) = "Empresa='" & LcEmpresa & "'"
Relatorio.Formulas(2) = "Endereco='" & LcEndereco & "'"
Relatorio.Formulas(3) = "Fone='" & LcFone & "'"

If impressora Then
   LcTipoSaida = 1
Else
   LcTipoSaida = 0
End If
'Relatorio.SelectionFormula = LcFormula
Relatorio.Destination = LcTipoSaida
Relatorio.PrintReport

If Relatorio.LastErrorNumber > 0 Then
   If Relatorio.LastErrorString <> "No Error" Then
     If Len(Trim(Relatorio.LastErrorString)) <> 0 Then
        MsgBox Relatorio.LastErrorString
     End If
   End If
End If

End Sub

Private Sub Convenio_Click()
Call Exibicao(1)
titulo1(0).Caption = "Convênio"
Titulo2.Visible = False
Txt(1).Visible = False
LcCampo = "Convenio"
LcDetalhe = "'"
End Sub

Private Sub Dataf_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then cmdOK.SetFocus

End Sub

Private Sub Datai_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Dataf.SetFocus

End Sub

Private Sub Endereço_Click()
Call Exibicao(1)
titulo1(0).Caption = "Codigo"
Titulo2.Visible = False
Txt(0).Visible = True
Txt(1).Visible = False
Datai.Visible = False
Dataf.Visible = False
LcCampo = "Produto"
LcDetalhe = "'"
End Sub

Private Sub Form_Load()
On Error Resume Next
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2

LcEscolha = "Igual"
Call Exibicao(1)
BuscaExpressao
LcCampo = "Descricao"
LcDetalhe = "'"
End Sub



Private Sub Impressora_Click()
On Error Resume Next
titulo1(1).Visible = True
copias.Visible = True
End Sub

Private Sub nome_Click()
Call Exibicao(1)
titulo1(0).Caption = "Produto"
Titulo2.Visible = False
Txt(0).Visible = True
Txt(1).Visible = False
Datai.Visible = False
Dataf.Visible = False
LcCampo = "Descricao"
LcDetalhe = "'"
End Sub

Private Sub Opt1_Click()
Escolha
BuscaExpressao
End Sub

Private Sub Opt2_Click()
Escolha
BuscaExpressao
End Sub


Private Sub Opt3_Click()
Escolha
BuscaExpressao
End Sub

Private Sub Opt4_Click()
Escolha
BuscaExpressao
End Sub

Private Sub Opt5_Click()
Escolha
BuscaExpressao
End Sub

Private Sub Opt6_Click()
Escolha
BuscaExpressao
End Sub

Private Sub txt_GotFocus(Index As Integer)
On Error Resume Next
Txt(0).Text = ""
Txt(1).Text = ""
End Sub

Private Sub Vencimento_Click()
Call Exibicao(2)
titulo1(0).Caption = "Vencimento"
Titulo2.Visible = False
Txt(1).Visible = False
LcCampo = "DiaVencimento"
LcDetalhe = "'"
End Sub
Function Escolha()
LcEscolha = Screen.ActiveControl.Caption

End Function
Function BuscaExpressao()
On Error Resume Next
If Len(Trim(LcCap1)) <> 0 Then
   titulo1(0).Caption = LcCap1
End If
Titulo2.Visible = False
Txt(1).Visible = False
Select Case LcEscolha
       Case Is = "Igual"
           LcExpressao = "="
        
       Case Is = "Menor"
            LcExpressao = "<"
       Case Is = "Maior"
            LcExpressao = ">"
       Case Is = "Menor Igual"
            LcExpressao = "<="
       Case Is = "Maior Igual"
            LcExpressao = ">="
       Case Is = "Entre"
            LcExpressao = "BETWEEN"
            LcCap1 = titulo1(0).Caption
            titulo1(0).Caption = "Início"
            Titulo2.Caption = "Fim"
            Titulo2.Visible = True
            Txt(1).Visible = True
       Case Is = "Iniciado Por"
            LcExpressao = "Like"
       Case Is = "Que Tenha"
            LcExpressao = "Like"
End Select

End Function

Private Sub Vídeo_Click()
On Error Resume Next
titulo1(1).Visible = False
copias.Visible = False
End Sub
