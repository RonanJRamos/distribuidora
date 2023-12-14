VERSION 5.00
Begin VB.Form ConfiguraSomaEntrada 
   BackColor       =   &H00D8C5B6&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configurações"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   6045
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox DespAce 
      BackColor       =   &H00D8C5B6&
      Caption         =   "Despesas Acessorias"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   5415
   End
   Begin VB.CheckBox NaoTributado 
      BackColor       =   &H00D8C5B6&
      Caption         =   "Serviço não Tributado"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   5175
   End
   Begin VB.CheckBox ValorIcmsSubst 
      BackColor       =   &H00D8C5B6&
      Caption         =   "Valor do Icms Substituição"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   5055
   End
   Begin VB.CheckBox Complementar 
      BackColor       =   &H00D8C5B6&
      Caption         =   "Complementar"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   5175
   End
   Begin VB.CheckBox PIS_COFINS 
      BackColor       =   &H00D8C5B6&
      Caption         =   "PIS/COFINS"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   5175
   End
   Begin VB.CheckBox Seguro 
      BackColor       =   &H00D8C5B6&
      Caption         =   "Seguro"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   5055
   End
   Begin VB.CheckBox Frete 
      BackColor       =   &H00D8C5B6&
      Caption         =   "Frete"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   4935
   End
   Begin VB.CommandButton CmdFechar 
      Caption         =   "Fechar"
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CommandButton CmdSalvar 
      Caption         =   "Salvar"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   3960
      Width           =   2175
   End
End
Attribute VB_Name = "ConfiguraSomaEntrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private LcCaminho As String

Private Sub CmdSalvar_Click()
On Error Resume Next

GravaIni "Soma", "Frete", Frete.Value, LcCaminho
GravaIni "Soma", "Seguro", Seguro.Value, LcCaminho
GravaIni "Soma", "PIS_COFINS", PIS_COFINS.Value, LcCaminho
GravaIni "Soma", "Complementar", Complementar.Value, LcCaminho
GravaIni "Soma", "ValorIcmsSubst", ValorIcmsSubst.Value, LcCaminho
GravaIni "Soma", "NaoTributado", NaoTributado.Value, LcCaminho
GravaIni "Soma", "DespAce", DespAce.Value, LcCaminho

MsgBox "Configuração salva.", 64, "Aviso"

End Sub

Private Sub Form_Load()
On Error Resume Next
GlNomeMaquina = ""
LcCaminho = App.Path & "\configMeiaFolha.ini"

Frete.Value = IIf(Len(LeIni("Soma", "Frete", LcCaminho)) = 0, 0, LeIni("Soma", "Frete", LcCaminho))
Seguro.Value = IIf(Len(LeIni("Soma", "Seguro", LcCaminho)) = 0, 0, LeIni("Soma", "Seguro", LcCaminho))
PIS_COFINS.Value = IIf(Len(LeIni("Soma", "PIS_COFINS", LcCaminho)) = 0, 0, LeIni("Soma", "PIS_COFINS", LcCaminho))
Complementar.Value = IIf(Len(LeIni("Soma", "Complementar", LcCaminho)) = 0, 0, LeIni("Soma", "Complementar", LcCaminho))
ValorIcmsSubst.Value = IIf(Len(LeIni("Soma", "ValorIcmsSubst", LcCaminho)) = 0, 0, LeIni("Soma", "ValorIcmsSubst", LcCaminho))
NaoTributado.Value = IIf(Len(LeIni("Soma", "NaoTributado", LcCaminho)) = 0, 0, LeIni("Soma", "NaoTributado", LcCaminho))
DespAce.Value = IIf(Len(LeIni("Soma", "DespAce", LcCaminho)) = 0, 0, LeIni("Soma", "DespAce", LcCaminho))



End Sub
