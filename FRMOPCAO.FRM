VERSION 5.00
Begin VB.Form FrmOpcao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Opções do Sistema"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8415
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   8415
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Reserva 
      Height          =   375
      Index           =   4
      Left            =   6360
      TabIndex        =   24
      Top             =   5280
      Width           =   1815
   End
   Begin VB.TextBox txt 
      Height          =   1455
      Index           =   1
      Left            =   4200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   22
      Top             =   3600
      Width           =   3975
   End
   Begin VB.TextBox Reserva 
      Height          =   375
      Index           =   3
      Left            =   5760
      TabIndex        =   20
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox Reserva 
      Height          =   375
      Index           =   2
      Left            =   5760
      TabIndex        =   18
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox txt 
      Height          =   1455
      Index           =   0
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Top             =   3600
      Width           =   3975
   End
   Begin VB.TextBox Reserva 
      Height          =   375
      Index           =   1
      Left            =   5760
      TabIndex        =   14
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox Reserva 
      Height          =   375
      Index           =   0
      Left            =   5760
      TabIndex        =   12
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Fecha"
      Height          =   495
      Left            =   1800
      TabIndex        =   11
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tipo Devolução de Fitas"
      Height          =   2895
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   2535
      Begin VB.OptionButton Ultima 
         Caption         =   "Iguala  a Última"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2040
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.OptionButton Primeira 
         Caption         =   "Iguala a Primeira"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   2055
      End
      Begin VB.OptionButton PorFita 
         Caption         =   "Por Fita"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Impressão em"
      Height          =   1095
      Left            =   2760
      TabIndex        =   1
      Top             =   240
      Width           =   2535
      Begin VB.OptionButton oitenta 
         Caption         =   "80 Colunas"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton quarenta 
         Caption         =   "40 Colunas"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Calcula Código Mercadoria"
      Height          =   1095
      Left            =   2760
      TabIndex        =   0
      Top             =   2040
      Width           =   2535
      Begin VB.OptionButton CodMercNao 
         Caption         =   "Não"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1935
      End
      Begin VB.OptionButton CodigoMercSim 
         Caption         =   "Sim"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   2055
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Salto de Linhas Fim da Imp."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   3480
      TabIndex        =   25
      Top             =   5400
      Width           =   2760
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Mensagem Felicitação"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   4200
      TabIndex        =   23
      Top             =   3360
      Width           =   2370
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Dias Dev. Lanc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   5760
      TabIndex        =   21
      Top             =   1680
      Width           =   1590
   End
   Begin VB.Line Line4 
      X1              =   7920
      X2              =   7920
      Y1              =   240
      Y2              =   3120
   End
   Begin VB.Line Line3 
      X1              =   5400
      X2              =   7920
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   5400
      X2              =   7920
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   5400
      X2              =   5400
      Y1              =   240
      Y2              =   3120
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Taxa Rebobinação"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   5760
      TabIndex        =   19
      Top             =   2400
      Width           =   2010
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Mensagem Conprovante"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   120
      TabIndex        =   17
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Dias Devolução"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   5760
      TabIndex        =   15
      Top             =   960
      Width           =   1680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Dias de Reserva"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   5760
      TabIndex        =   13
      Top             =   360
      Width           =   1770
   End
End
Attribute VB_Name = "FrmOpcao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CodigoMercSim_Click()
On Error Resume Next
GlCalculaProduto = True
End Sub

Private Sub CodMercNao_Click()
On Error Resume Next
GlCalculaProduto = False
End Sub

Private Sub Command1_Click()

gravaopcoes
Unload Me
End Sub

Private Sub Command2_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next

Me.Height = 6330
Me.Width = 8505
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
VerificaOpcoes
End Sub
Function VerificaOpcoes()
On Error Resume Next
If GlImpressao Then
   quarenta = True
   oitenta = False
Else
   quarenta = False
   oitenta = True
End If

If GlCalculaProduto Then
   CodigoMercSim = True
   CodMercNao = False
Else
   CodigoMercSim = False
   CodMercNao = True
End If

Select Case GLtipoDevolucao
       Case Is = 1
            PorFita = True
            Primeira = False
            Ultima = False
            
       Case Is = 2
            PorFita = False
            Primeira = True
            Ultima = False

       Case Is = 3
            PorFita = False
            Primeira = False
            Ultima = True

End Select
Reserva(0).Text = GlDiaReserva
Reserva(1).Text = GlDiasDevolucao
Reserva(2).Text = GlaxaReb
Reserva(3).Text = GlDiasDevolucaoReserva
Reserva(4).Text = GlMsg3
Open App.Path & "\Msg.txt" For Input As #Fnum    ' Open file for output.
Input #Fnum, GlMsg
Close #Fnum
Fnum = FreeFile
Open App.Path & "\Felic.txt" For Input As #Fnum    ' Open file for output.
Input #Fnum, GlMsg1
Close #Fnum
txt(0).Text = GlMsg
txt(1).Text = GlMsg1
End Function
Function gravaopcoes()
On Error Resume Next
Dim a As Integer, Fnum As Integer
Call AbreBanco(opcao)

a = RsAtual!CalculaCodigoProduto
If Err <> 0 Then
   RsAtual.AddNew
Else
   RsAtual.Edit
End If
RsAtual!CalculaCodigoProduto = GlCalculaProduto
RsAtual!ImprimeQuarentaColunas = GlImpressao
RsAtual!TipodevolucaoFita = GLtipoDevolucao
RsAtual!DiasReserva = GlDiaReserva
RsAtual!DiasDevolucao = GlDiasDevolucao
RsAtual!Rebobinacao = GlaxaReb
RsAtual!msg3 = GlMsg3

RsAtual!DiasDevolucaoLancamento = GlDiasDevolucaoReserva
RsAtual.Update
Call FechaBanco
Fnum = FreeFile
Open App.Path & "\Msg.txt" For Output As #Fnum    ' Open file for output.
Write #Fnum, txt(0).Text
Close #Fnum
Fnum = FreeFile
Open App.Path & "\Felic.txt" For Output As #Fnum    ' Open file for output.
Write #Fnum, txt(1).Text
Close #Fnum
End Function

Private Sub oitenta_Click()
On Error Resume Next
GlImpressao = False
End Sub

Private Sub PorFita_Click()
On Error Resume Next

GLtipoDevolucao = 1
End Sub

Private Sub Primeira_Click()
On Error Resume Next
GLtipoDevolucao = 2
End Sub

Private Sub quarenta_Click()
On Error Resume Next
GlImpressao = True
End Sub



Private Sub Reserva_Change(Index As Integer)

If Index = 0 Then GlDiaReserva = Val(Reserva(0).Text)
If Index = 1 Then GlDiasDevolucao = Val(Reserva(1).Text)
If Index = 2 Then GlaxaReb = CCur(Reserva(2).Text)
If Index = 3 Then GlDiasDevolucaoReserva = CCur(Reserva(3).Text)
If Index = 4 Then GlMsg3 = Val(Reserva(4).Text)
End Sub

Private Sub Txt_Change(Index As Integer)
On Error Resume Next
Select Case Index
   Case Is = 0
     GlMsg = txt(Index).Text
   Case Is = 1
     GlMsg1 = txt(Index).Text
   Case Is = 2
     GlMsg2 = txt(Index).Text
   Case Is = 3
    GlMsg3 = txt(Index).Text
End Select

End Sub

Private Sub Ultima_Click()
On Error Resume Next
GLtipoDevolucao = 3
End Sub
