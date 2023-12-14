VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form FrmRelClientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Clientes"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8850
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   8850
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport Relatorio 
      Left            =   6240
      Top             =   4560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
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
      TabIndex        =   14
      Text            =   "1"
      Top             =   2400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton CmdFechar 
      Caption         =   "&Fechar"
      Height          =   615
      Left            =   5280
      TabIndex        =   18
      Top             =   3600
      Width           =   2415
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      Height          =   615
      Left            =   5280
      TabIndex        =   17
      Top             =   2880
      Width           =   2415
   End
   Begin VB.TextBox Txt 
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   16
      Top             =   5280
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.TextBox Txt 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   4560
      Width           =   5055
   End
   Begin VB.Frame Frame3 
      Caption         =   "Saída"
      Height          =   1575
      Left            =   5280
      TabIndex        =   24
      Top             =   360
      Width           =   2415
      Begin VB.OptionButton Impressora 
         Caption         =   "Impressora"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton Vídeo 
         Caption         =   "Vídeo"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo Comparação"
      Height          =   3855
      Left            =   2760
      TabIndex        =   23
      Top             =   360
      Width           =   2415
      Begin VB.OptionButton Opt6 
         Caption         =   "Iniciado Por"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   2640
         Width           =   1455
      End
      Begin VB.OptionButton Opt5 
         Caption         =   "Menor Igual"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   2160
         Width           =   1695
      End
      Begin VB.OptionButton Opt4 
         Caption         =   "Maior Igual"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1800
         Width           =   1575
      End
      Begin VB.OptionButton Opt3 
         Caption         =   "Maior"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   1695
      End
      Begin VB.OptionButton Opt2 
         Caption         =   "Menor"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   1815
      End
      Begin VB.OptionButton Opt1 
         Caption         =   "Igual"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filtrar"
      Height          =   3855
      Left            =   120
      TabIndex        =   22
      Top             =   360
      Width           =   2535
      Begin VB.OptionButton Vencimento 
         Caption         =   "Vencimento"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   2835
         Width           =   1935
      End
      Begin VB.OptionButton Convenio 
         Caption         =   "Convenio (Nome)"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   2340
         Width           =   1815
      End
      Begin VB.OptionButton Cidade 
         Caption         =   "Cidade"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1845
         Width           =   2055
      End
      Begin VB.OptionButton Bairro 
         Caption         =   "Bairro"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   1350
         Width           =   1935
      End
      Begin VB.OptionButton Endereço 
         Caption         =   "Endereço"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   855
         Width           =   2175
      End
      Begin VB.OptionButton Nome 
         Caption         =   "Nome"
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Value           =   -1  'True
         Width           =   1935
      End
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
      TabIndex        =   21
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
      Left            =   120
      TabIndex        =   20
      Top             =   5040
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
      TabIndex        =   19
      Top             =   4200
      Width           =   705
   End
End
Attribute VB_Name = "FrmRelClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private LcEscolha, LcDetalhe, LcCampo, LcExpressao, LcCap1 As String

Private Sub Atraso_Click()
Call Exibicao(2)
titulo1(0).Caption = "Dia Limite"
Titulo2.Visible = False
txt(1).Visible = False
LcCampo = "DiaVencimento"
LcDetalhe = ""
End Sub

Private Sub Bairro_Click()
Call Exibicao(1)
titulo1(0).Caption = "Bairro"
Titulo2.Visible = False
txt(1).Visible = False
LcCampo = "Bairro"
LcDetalhe = "'"
End Sub

Private Sub Cidade_Click()
Call Exibicao(1)
titulo1(0).Caption = "Cidade"
Titulo2.Visible = False
txt(1).Visible = False
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
On Error Resume Next
Dim RsEmpresa As Recordset
Dim a, item, LcResposta As Long
Dim LcCriterio, LcEmpresa, LcEndereco, LcFone As String

Set Dbbase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
Set RsEmpresa = Dbbase.OpenRecordset("Empresa", dbOpenDynaset)

LcEmpresa = RsEmpresa!RazaoDaEmpresa
LcEndereco = RTrim(RsEmpresa!EnderecoDaEmpresa) & RTrim(RsEmpresa!NumeroDaEmpresa) & " Bairro: " & RsEmpresa!Bairro & "  Cidade: " & RsEmpresa!cidade
LcFone = "Fone: " & RsEmpresa!fone
If Not IsNull(RsEmpresa!Fax) Then
   LcFone = LcFone & " Fax:  " & RsEmpresa!Fax
End If
If Err <> 0 Then
 LcEmpresa = ""
 LcFone = ""
 LcEndereco = ""
End If
RsEmpresa.Close

If Not Atraso Then
   Select Case LcExpressao
       Case Is = "BETWEEN"
          LcFormula = "{CLientes." & LcCampo & "} " & LcExpressao & " " & LcDetalhe & txt(0).Text & LcDetalhe & " And " & LcDetalhe & txt(1).Text & LcDetalhe
       Case Is = "Like"
          If LcEscolha = "Iniciado Por" Then
             LcFormula = "{CLientes." & LcCampo & "} " & LcExpressao & " " & LcDetalhe & UCase(txt(0).Text) & "%" & LcDetalhe _
             & " or {CLientes." & LcCampo & "} " & LcExpressao & " " & LcDetalhe & LCase(txt(0).Text) & "*" & LcDetalhe
          Else
             LcFormula = "{CLientes." & LcCampo & "} " & LcExpressao & " " & LcDetalhe & "*" & UCase(txt(0).Text) & "%" & LcDetalhe _
             & " or {CLientes." & LcCampo & "} " & LcExpressao & " " & LcDetalhe & "*" & LCase(txt(0).Text) & "%" & LcDetalhe
          End If
       Case Else
           LcFormula = "{CLientes." & LcCampo & "} " & LcExpressao & " " & LcDetalhe & UCase(txt(0).Text) & LcDetalhe _
            & " Or {CLientes." & LcCampo & "} " & LcExpressao & " " & LcDetalhe & LCase(txt(0).Text) & LcDetalhe
   End Select
Else
   LcFormula = "{CLientes." & LcCampo & "} <= " & Val((txt(0).Text)) & " And {CLientes.ValorDevido} >0"
End If

'MsgBox LcFormula
Relatorio.DataFiles(0) = GLBase

Relatorio.ReportFileName = App.Path & "\cliente.rpt"

Relatorio.SelectionFormula = LcFormula
Relatorio.SortFields(0) = "+{CLientes." & LcCampo & "}"
Relatorio.SortFields(1) = "+{CLientes." & LcCampo & "}"
Relatorio.CopiesToPrinter = Val(Copias.Text)


Relatorio.WindowTop = 50
Relatorio.WindowWidth = 700
Relatorio.WindowLeft = 50
Relatorio.WindowHeight = 500
Relatorio.WindowTitle = "Relatório de Clientes"

Relatorio.Formulas(0) = "titulo='Relátorio de Clientes'"
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
Relatorio.Password = "muralha"
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
txt(1).Visible = False
LcCampo = "Convenio"
LcDetalhe = "'"
End Sub

Private Sub Endereço_Click()
Call Exibicao(1)
titulo1(0).Caption = "Endereço"
Titulo2.Visible = False
txt(1).Visible = False
LcCampo = "Rua"
LcDetalhe = "'"
End Sub

Private Sub Form_Load()
On Error Resume Next
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2

LcEscolha = "Igual"
Call Exibicao(1)
BuscaExpressao
LcCampo = "Nome"
LcDetalhe = "'"
End Sub



Private Sub Impressora_Click()
On Error Resume Next
titulo1(1).Visible = True
Copias.Visible = True
End Sub

Private Sub Nome_Click()
Call Exibicao(1)
titulo1(0).Caption = "Nome"
Titulo2.Visible = False
txt(1).Visible = False
LcCampo = "Nome"
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
txt(0).Text = ""
txt(1).Text = ""
End Sub

Private Sub Vencimento_Click()
Call Exibicao(2)
titulo1(0).Caption = "Vencimento"
Titulo2.Visible = False
txt(1).Visible = False
LcCampo = "DiaVencimento"
LcDetalhe = ""
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
txt(1).Visible = False
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
            txt(1).Visible = True
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
