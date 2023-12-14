VERSION 5.00
Begin VB.Form FrmAviso 
   BackColor       =   &H00C5FEE1&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Erro No Acesso Ao Sistema"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9615
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   9615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Fechar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3480
      TabIndex        =   2
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Label Terceira 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   2040
      Width           =   9375
   End
   Begin VB.Label Segunda 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   9375
   End
   Begin VB.Label erro 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   9405
   End
End
Attribute VB_Name = "FrmAviso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private a As Integer
Private Sub Command1_Click()
On Error Resume Next
End
End Sub

Private Sub Form_Load()
On Error Resume Next
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
Me.FontName = "Arial"
Me.FontSize = 14
LcMsg1 = "Entre Em Contato Com o Suporte Técnico. (31) 3625-2906"
If Not GlErroProt Then
   err.Number = GlCodigoProtecao
   LcMsg = "Erro  Número :" & GlCodigoProtecao & " " & err.Description
   LcSegunda = True
Else
   Select Case GlCodigoProtecao
     Case Is = 1
       LcMsg = "Disco Protegido Contra Gravação..."
     Case Is = 2
       LcMsg = "Drive Não Pronto..."
     Case Is = 3
      LcMsg = "Programa Não Autorizado a Operar Nesta Máquina..."
      LcSegunda = True
   Case Is = 4
      LcMsg = "O Disco Inserido Não é o Disco Original..."
      LcSegunda = True
   Case Is = 5
      LcMsg = "Data de Demostração Encerrada..."
      LcSegunda = True
   Case Is = 6
      LcMsg = "A Data do Sistema Não Confere com a Última Execusão do Sistema."
      LcSegunda = False
   Case Is = 7
      LcMsg = "Limite de Acesso a Rede Atinguido..."
      LcSegunda = True
   Case Is = 8
      LcMsg = "Erro de Acesso ao disco.."
      LcSegunda = True
   Case Is = 9
      LcMsg = "Sistema Não Autorizado a Operar nesta Máquina..."
      LcSegunda = True
   Case Is = 12
      LcMsg = "Número de Instações Esgotadas..."
      LcSegunda = True
   Case Is = 13
      LcMsg = "Sistema Não Autorizado a Operar Nesta Máquina.."
      LcSegunda = True
  Case Is = 15
      LcMsg = "Sistema Não Autorizado a Operar Nesta Máquina.."
      LcSegunda = True
  Case Is = 18
      LcMsg = "Sistema Não Autorizado a Operar Nesta Máquina.."
      LcSegunda = True
  Case Is = 19
     LcMsg = "Disco de Instalação Inválido..."
     LcSegunda = True
  Case Is = 20
     LcMsg = "Disco de Instalação Inválido..."
     LcSegunda = True
  Case Else
     LcMsg = "Erro Desconhecido Número :" & GlCodigoProtecao
     LcSegunda = True
   
 End Select
End If

Erro.Caption = LcMsg
Me.Caption = "Nº do Erro: " & GlCodigoProtecao
If LcSegunda Then
   Segunda.Caption = LcMsg1
Else
   Segunda.Caption = ""
End If
End Sub
