VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Empresa 
   BackColor       =   &H00CAE1A2&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dados da Empresa"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10170
   Icon            =   "Empresa.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   10170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdFechar 
      Caption         =   "&Fechar"
      Height          =   495
      Left            =   8640
      TabIndex        =   37
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton CmdSalvar 
      Caption         =   "&Salvar"
      Height          =   495
      Left            =   8640
      TabIndex        =   36
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox Modelo 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   7560
      TabIndex        =   34
      Top             =   3960
      Width           =   855
   End
   Begin VB.TextBox Responsavel 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1320
      TabIndex        =   33
      Top             =   3960
      Width           =   4935
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00CAE1A2&
      Caption         =   "FINALIDADES DA APRESENTAÇÃO DO ARQUIVO MAGNÉTICO"
      Height          =   735
      Left            =   120
      TabIndex        =   31
      Top             =   6120
      Width           =   8295
      Begin VB.ComboBox Finalidade 
         Height          =   315
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   7935
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CAE1A2&
      Caption         =   "Código da identificação da natureza das operações informadas"
      Height          =   735
      Left            =   120
      TabIndex        =   29
      Top             =   5280
      Width           =   8295
      Begin VB.ComboBox CodigoIdentifciacaoNatureza 
         Height          =   315
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   7935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CAE1A2&
      Caption         =   "CÓDIGO DE IDENTIFICAÇÃO DA ESTRUTURA DO ARQUIVO MAGNÉTICO "
      Height          =   735
      Left            =   120
      TabIndex        =   27
      Top             =   4560
      Width           =   8295
      Begin VB.ComboBox CodigoEstrutura 
         Height          =   315
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   7935
      End
   End
   Begin MSMask.MaskEdBox Cnpj 
      Height          =   375
      Left            =   1320
      TabIndex        =   22
      Top             =   3000
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   18
      Mask            =   "99.999.999/9999-99"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Cep 
      Height          =   375
      Left            =   1320
      TabIndex        =   19
      Top             =   2520
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   8
      Mask            =   "99999-99"
      PromptChar      =   " "
   End
   Begin VB.TextBox Email 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1320
      TabIndex        =   10
      Top             =   3480
      Width           =   4935
   End
   Begin VB.TextBox Inscricao 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   4920
      TabIndex        =   9
      Top             =   3000
      Width           =   3495
   End
   Begin VB.TextBox Fax 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   6360
      TabIndex        =   8
      Top             =   2520
      Width           =   2055
   End
   Begin VB.TextBox Fone 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox Estado 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   7080
      TabIndex        =   6
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox Cidade 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   2040
      Width           =   4935
   End
   Begin VB.TextBox Bairro 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   1560
      Width           =   4935
   End
   Begin VB.TextBox Complemento 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   1080
      Width           =   4935
   End
   Begin VB.TextBox Numero 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   7080
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox Endereco 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   600
      Width           =   4935
   End
   Begin VB.TextBox Razao 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   7095
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Modelo Nota"
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
      Left            =   6360
      TabIndex        =   35
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Responsável"
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
      TabIndex        =   26
      Top             =   4080
      Width           =   1110
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail"
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
      TabIndex        =   25
      Top             =   3600
      Width           =   525
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inscrição Estadual"
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
      Left            =   3120
      TabIndex        =   24
      Top             =   3120
      Width           =   1590
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cnpj"
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
      TabIndex        =   23
      Top             =   3120
      Width           =   390
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fax"
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
      Left            =   5640
      TabIndex        =   21
      Top             =   2640
      Width           =   315
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fone"
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
      Left            =   2520
      TabIndex        =   20
      Top             =   2640
      Width           =   435
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cep"
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
      TabIndex        =   18
      Top             =   2640
      Width           =   345
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Uf"
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
      Left            =   6840
      TabIndex        =   17
      Top             =   2160
      Width           =   210
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cidade"
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
      TabIndex        =   16
      Top             =   2160
      Width           =   600
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bairro"
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
      TabIndex        =   15
      Top             =   1680
      Width           =   510
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Complemento"
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
      TabIndex        =   14
      Top             =   1200
      Width           =   1140
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº"
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
      Left            =   6840
      TabIndex        =   13
      Top             =   690
      Width           =   225
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Endereço"
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
      TabIndex        =   12
      Top             =   720
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Razão Social"
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
      TabIndex        =   11
      Top             =   240
      Width           =   1140
   End
End
Attribute VB_Name = "Empresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()

End Sub

Private Sub CmdFechar_Click()
On Error Resume Next
Unload Me
End Sub
Sub BuscaDados()
On Error Resume Next
Dim Rs As Recordset
Dim db As Database
AbreBase
'Set db = OpenRecordset(GLBase)
Set Rs = Dbbase.OpenRecordset("Select * from empresa")

If Not Rs.EOF Then
   Razao.Text = Rs!Razao & ""
   Endereco.Text = Rs!Endereco & ""
   Numero.Text = Rs!Numero & ""
   Complemento.Text = Rs!Complemento & ""
   Bairro.Text = Rs!Bairro & ""
   Cidade.Text = Rs!Cidade & ""
   Estado.Text = Rs!Estado & ""
   Cep.Text = Rs!Cep & ""
   Fone.Text = Rs!Fone & ""
   Fax.Text = Rs!Fax & ""
   Cnpj.Text = Rs!CGC & ""
   Inscricao.Text = Rs!inscricaoestadual & ""
   Email.Text = Rs!Email & ""
   Responsavel.Text = Rs!Responsavel & ""
   Modelo.Text = "01"
   '==> Busca os Codigos
   Dim codigoConv As Integer
   Dim CodigoIdent As Integer
   Dim CodFina As Integer
   CodFina = Rs!CodigofianlidadeArquivo
   CodigoIdent = Rs!CodigoNaturezaInformacao
   codigoConv = Rs!CodigoConvenio
   If Not IsNull(Rs!CodigoConvenio) Then CodigoEstrutura.ListIndex = codigoConv - 1
   If Not IsNull(Rs!CodigoNaturezaInformacao) Then CodigoIdentifciacaoNatureza.ListIndex = CodigoIdent - 1
   If Not IsNull(Rs!CodigofianlidadeArquivo) Then Finalidade.ListIndex = CodFina - 1
End If
Set Rs = Nothing

End Sub
Sub Salva()
On Error Resume Next
Dim Rs As Recordset
AbreBase
Set Rs = Dbbase.OpenRecordset("Select * from empresa")

If Not Rs.EOF Then
    Rs.Edit
Else
    Rs.AddNew
End If
Rs!Razao = Razao.Text
Rs!Endereco = Endereco.Text
Numero.Text = Rs!Numero & ""
Rs!Complemento = Complemento.Text
Rs!Bairro = Bairro.Text
Rs!Cidade = Cidade.Text
Rs!Estado = Estado.Text
Rs!Cep = Cep.Text
Rs!Fone = Fone.Text
Rs!Fax = Fax.Text
Rs!CGC = Cnpj.Text
Rs!inscricaoestadual = Inscricao.Text
Rs!Email = Email.Text
Rs!Responsavel = Responsavel.Text
Rs!Modelo = Modelo.Text
   '==> Busca os Codigos
Rs!CodigoConvenio = Mid(CodigoEstrutura.Text, 1, 1)
Rs!CodigoNaturezaInformacao = Mid(CodigoIdentifciacaoNatureza.Text, 1, 1)
Rs!CodigofianlidadeArquivo = Mid(Finalidade.Text, 1, 1)

Rs.Update

Set Rs = Nothing

End Sub
Private Sub Form_Load()
On Error Resume Next
CodigoEstrutura.AddItem "1- Estrutura conforme Convênio ICMS 57/95, na versão estabelecida pelo Convênio ICMS 31/99 e com as alterações promovidas até o Convênio ICMS 30/02."
CodigoEstrutura.AddItem "2- Estrutura conforme Convênio ICMS 57/95, na versão estabelecida pelo Convênio ICMS 69/02 e com as alterações promovidas pelo Convênio ICMS 142/02."
CodigoEstrutura.AddItem "3- Estrutura conforme Convênio ICMS 57/95, com as alterações promovidas pelo Convênio ICMS 76/03."


CodigoIdentifciacaoNatureza.AddItem "1- Interestaduais somente operações sujeitas ao regime de Substituição Tributária."
CodigoIdentifciacaoNatureza.AddItem "2- Interestaduais - operações com ou sem Substituição Tributária."
CodigoIdentifciacaoNatureza.AddItem "3- Totalidade das operações do informante."

Finalidade.AddItem "1- Normal."
Finalidade.AddItem "2- Retificação total de arquivo: substituição total de informações prestadas pelo contribuinte referentes a este período."
Finalidade.AddItem "3- Retificação aditiva de arquivo: acréscimo de informação não incluída em arquivos já apresentados."
Finalidade.AddItem "5- Desfazimento: arquivo de informação referente a operações/prestações não efetivadas."
BuscaDados
End Sub
