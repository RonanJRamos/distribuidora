VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form RelSaidaProdutos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatorio de Saída de Produtos"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7905
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   7905
   StartUpPosition =   3  'Windows Default
   Begin MSMask.MaskEdBox datai 
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   3360
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
      Left            =   5400
      TabIndex        =   11
      Text            =   "1"
      Top             =   2400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin Crystal.CrystalReport Relatorio 
      Left            =   6960
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.CommandButton CmdFechar 
      Caption         =   "&Fechar"
      Height          =   615
      Left            =   5400
      TabIndex        =   15
      Top             =   3960
      Width           =   2415
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      Height          =   615
      Left            =   5400
      TabIndex        =   14
      Top             =   3120
      Width           =   2415
   End
   Begin VB.TextBox Txt 
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   3720
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.TextBox Txt 
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   12
      Top             =   3720
      Width           =   5055
   End
   Begin VB.Frame Frame3 
      Caption         =   "Saída"
      Height          =   1575
      Left            =   5280
      TabIndex        =   21
      Top             =   360
      Width           =   2415
      Begin VB.OptionButton Impressora 
         Caption         =   "Impressora"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton Vídeo 
         Caption         =   "Vídeo"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo Comparação"
      Height          =   2535
      Left            =   2760
      TabIndex        =   20
      Top             =   360
      Width           =   2415
      Begin VB.OptionButton Opt6 
         Caption         =   "Iniciado Por"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   1920
         Width           =   1455
      End
      Begin VB.OptionButton Opt5 
         Caption         =   "Menor Igual"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   1800
         Width           =   1695
      End
      Begin VB.OptionButton Opt4 
         Caption         =   "Maior Igual"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1800
         Width           =   1575
      End
      Begin VB.OptionButton Opt3 
         Caption         =   "Maior"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   1695
      End
      Begin VB.OptionButton Opt2 
         Caption         =   "Menor"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   1815
      End
      Begin VB.OptionButton Opt1 
         Caption         =   "Igual"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filtrar"
      Height          =   2535
      Left            =   120
      TabIndex        =   19
      Top             =   360
      Width           =   2535
      Begin VB.OptionButton Bairro 
         Caption         =   "Período"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   1680
         Width           =   1935
      End
      Begin VB.OptionButton Endereço 
         Caption         =   "Código"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   1080
         Width           =   1455
      End
      Begin VB.OptionButton Nome 
         Caption         =   "Produto"
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Value           =   -1  'True
         Width           =   1935
      End
   End
   Begin MSMask.MaskEdBox Dataf 
      Height          =   375
      Left            =   3000
      TabIndex        =   23
      Top             =   3360
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
      Left            =   5400
      TabIndex        =   18
      Top             =   2040
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
      Top             =   3360
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
      Top             =   3000
      Width           =   705
   End
End
Attribute VB_Name = "RelSaidaProdutos"
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
datai.Visible = True
Dataf.Visible = True
datai.SetFocus
LcCampo = "DataSaida"
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
Dim LcCriterio, LcEmpresa, LcEndereco, LcFone, lctitulo As String


Set DbBase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
Set RsEmpresa = DbBase.OpenRecordset("Empresa", dbOpenDynaset)

LcEmpresa = RsEmpresa!RazaoDaEmpresa
LcEndereco = RTrim(RsEmpresa!EnderecoDaEmpresa) & RTrim(RsEmpresa!NumeroDaEmpresa) & " Bairro: " & RsEmpresa!Bairro & "  Cidade: " & RsEmpresa!Cidade
LcFone = "Fone: " & RsEmpresa!fone
If Not IsNull(RsEmpresa!Fax) Then
   LcFone = LcFone & " Fax: " & RsEmpresa!Fax
End If
If Err <> 0 Then
 LcEmpresa = ""
 LcFone = ""
 LcEndereco = ""
End If
RsEmpresa.Close
If LcCampo = "DataSaida" Then
         strData = CDate(Format(datai.Text, "dd/mm/yyyy"))
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
         LcFormula = "{SaidaProduto.DataSaida} >=" & LcChav1 & " And {SaidaProduto.datasaida} <=" & LcChav2
         lctitulo = "Saída de Produtos Período " & Format(datai.Text, "dd/mm/yyyy") & " Até " & Format(Dataf.Text, "dd/mm/yyyy")
Else
   lctitulo = "Saída de Produtos "
  Select Case LcExpressao
       Case Is = "BETWEEN"
          LcFormula = "{SaidaProduto." & LcCampo & "} " & LcExpressao & " " & LcDetalhe & Txt(0).Text & LcDetalhe & " And " & LcDetalhe & Txt(1).Text & LcDetalhe
       Case Is = "Like"
          If LcEscolha = "Iniciado Por" Then
             LcFormula = "{SaidaProduto." & LcCampo & "} " & LcExpressao & " " & LcDetalhe & UCase(Txt(0).Text) & "*" & LcDetalhe _
             & " or {SaidaProduto." & LcCampo & "} " & LcExpressao & " " & LcDetalhe & LCase(Txt(0).Text) & "*" & LcDetalhe
          Else
             LcFormula = "{SaidaProduto." & LcCampo & "} " & LcExpressao & " " & LcDetalhe & "*" & UCase(Txt(0).Text) & "*" & LcDetalhe _
             & " or {SaidaProduto." & LcCampo & "} " & LcExpressao & " " & LcDetalhe & "*" & LCase(Txt(0).Text) & "*" & LcDetalhe
          End If
       Case Else
           LcFormula = "{SaidaProduto." & LcCampo & "} " & LcExpressao & " " & LcDetalhe & UCase(Txt(0).Text) & LcDetalhe _
            & " Or {SaidaProduto." & LcCampo & "} " & LcExpressao & " " & LcDetalhe & LCase(Txt(0).Text) & LcDetalhe
  End Select
End If
' MsgBox LcFormula
Relatorio.DataFiles(0) = GLBase
Relatorio.ReportFileName = App.Path & "\SaidaProduto.rpt"
Relatorio.SelectionFormula = LcFormula

Relatorio.SortFields(0) = "+{SaidaProduto." & LcCampo & "}"
Relatorio.CopiesToPrinter = Val(Copias.Text)


Relatorio.WindowTop = 50
Relatorio.WindowWidth = 700
Relatorio.WindowLeft = 50
Relatorio.WindowHeight = 500
Relatorio.WindowTitle = "Saída de Produtos "

Relatorio.Formulas(0) = "titulo='" & lctitulo & "'"
Relatorio.Formulas(1) = "Empresa='" & LcEmpresa & "'"
Relatorio.Formulas(2) = "EnderecoEmpresa='" & LcEndereco & "'"
Relatorio.Formulas(3) = "Fone='" & LcFone & "'"

If Impressora Then
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
If KeyCode = 13 Then CmdOk.SetFocus

End Sub

Private Sub datai_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Dataf.SetFocus

End Sub

Private Sub Endereço_Click()
Call Exibicao(1)
titulo1(0).Caption = "Codigo"
Titulo2.Visible = False
Txt(0).Visible = True
Txt(1).Visible = False
datai.Visible = False
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
Copias.Visible = True
End Sub

Private Sub Nome_Click()
Call Exibicao(1)
titulo1(0).Caption = "Produto"
Titulo2.Visible = False
Txt(0).Visible = True
Txt(1).Visible = False
datai.Visible = False
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
Copias.Visible = False
End Sub
