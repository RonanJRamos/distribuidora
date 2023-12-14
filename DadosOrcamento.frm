VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form DadosOrcamento 
   BackColor       =   &H00CAE1A2&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dados Finais do Orçamento"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10440
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   10440
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Visualizar 
      BackColor       =   &H00CAE1A2&
      Caption         =   "Visualizar"
      Height          =   255
      Left            =   8760
      TabIndex        =   48
      Top             =   1800
      Width           =   1575
   End
   Begin MSMask.MaskEdBox PrevCommisao 
      Height          =   375
      Left            =   7080
      TabIndex        =   13
      Top             =   1680
      Width           =   1455
      _ExtentX        =   2566
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
   Begin VB.TextBox TipoPag 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Dados Transp F3"
      Height          =   375
      Left            =   8760
      TabIndex        =   10
      Top             =   120
      Width           =   1575
   End
   Begin VB.ComboBox TipoMonetario 
      Height          =   315
      Left            =   4080
      TabIndex        =   1
      Text            =   "Nenhum"
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton Imprime 
      Caption         =   "&Imprimir F2"
      Height          =   375
      Left            =   8760
      TabIndex        =   11
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar F10"
      Height          =   375
      Left            =   8760
      TabIndex        =   12
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox quantidade 
      Height          =   375
      Left            =   3480
      TabIndex        =   9
      Top             =   1680
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   11
      Left            =   6720
      TabIndex        =   2
      Text            =   "0"
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   12
      Left            =   6120
      MaxLength       =   12
      TabIndex        =   22
      Top             =   4680
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   10
      Left            =   1920
      MaxLength       =   30
      TabIndex        =   24
      Top             =   5040
      Width           =   2535
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   9
      Left            =   1920
      MaxLength       =   55
      TabIndex        =   27
      Top             =   6480
      Width           =   5655
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   8
      Left            =   1920
      MaxLength       =   55
      TabIndex        =   26
      Top             =   6165
      Width           =   5655
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   7
      Left            =   1920
      MaxLength       =   55
      TabIndex        =   25
      Top             =   5880
      Width           =   5655
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   6
      Left            =   1920
      MaxLength       =   25
      TabIndex        =   21
      Top             =   4680
      Width           =   3135
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   5
      Left            =   8280
      MaxLength       =   2
      TabIndex        =   23
      Top             =   4680
      Width           =   735
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   4
      Left            =   1920
      MaxLength       =   20
      TabIndex        =   20
      Top             =   4320
      Width           =   7095
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   3
      Left            =   1920
      MaxLength       =   50
      TabIndex        =   19
      Top             =   3960
      Width           =   7095
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   2
      Left            =   6120
      MaxLength       =   18
      TabIndex        =   18
      Top             =   3120
      Width           =   2895
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   1
      Left            =   3720
      MaxLength       =   2
      TabIndex        =   17
      Top             =   3120
      Width           =   735
   End
   Begin MSMask.MaskEdBox Placa 
      Height          =   375
      Left            =   1920
      TabIndex        =   16
      Top             =   3000
      Width           =   1215
      _ExtentX        =   2143
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
      Format          =   "AAA-9999"
      Mask            =   "AAA-9999"
      PromptChar      =   " "
   End
   Begin VB.ComboBox Tipo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "DadosOrcamento.frx":0000
      Left            =   7440
      List            =   "DadosOrcamento.frx":000A
      TabIndex        =   15
      Text            =   "1- CIF"
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   0
      Left            =   1920
      MaxLength       =   30
      TabIndex        =   14
      Top             =   2640
      Width           =   4815
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   1680
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Vencimento 
      Height          =   375
      Index           =   0
      Left            =   1920
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
      _ExtentX        =   1931
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
   Begin MSMask.MaskEdBox Vencimento 
      Height          =   375
      Index           =   1
      Left            =   3120
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
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
   Begin MSMask.MaskEdBox Vencimento 
      Height          =   375
      Index           =   2
      Left            =   4320
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
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
   Begin MSMask.MaskEdBox Vencimento 
      Height          =   375
      Index           =   3
      Left            =   5520
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
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
   Begin MSMask.MaskEdBox Vencimento 
      Height          =   375
      Index           =   4
      Left            =   6720
      TabIndex        =   7
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
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
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Previsão Comissão"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   47
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Line Line4 
      X1              =   10440
      X2              =   0
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Forma Pag."
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
      Index           =   10
      Left            =   2880
      TabIndex        =   46
      Top             =   480
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   8640
      X2              =   8640
      Y1              =   0
      Y2              =   2400
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   7800
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cond.Pag"
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
      Index           =   12
      Left            =   120
      TabIndex        =   45
      Top             =   480
      Width           =   1035
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valores"
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
      Index           =   14
      Left            =   120
      TabIndex        =   44
      Top             =   1680
      Width           =   825
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimentos"
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
      Index           =   13
      Left            =   120
      TabIndex        =   43
      Top             =   1200
      Width           =   1350
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Forma Pagamento"
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
      Height          =   330
      Index           =   15
      Left            =   2640
      TabIndex        =   42
      Top             =   0
      Width           =   2550
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dias"
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
      Index           =   17
      Left            =   6240
      TabIndex        =   41
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C.E.P.:"
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
      Index           =   18
      Left            =   5160
      TabIndex        =   40
      Top             =   4680
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dados Para Entrega"
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
      Height          =   330
      Index           =   16
      Left            =   2640
      TabIndex        =   39
      Top             =   3480
      Width           =   2790
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fone"
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
      Index           =   11
      Left            =   120
      TabIndex        =   38
      Top             =   4920
      Width           =   540
   End
   Begin VB.Line Line2 
      X1              =   7800
      X2              =   8640
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Informações Complementares"
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
      Height          =   330
      Index           =   9
      Left            =   2640
      TabIndex        =   37
      Top             =   5400
      Width           =   4185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cidade"
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
      Index           =   8
      Left            =   120
      TabIndex        =   36
      Top             =   4680
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UF"
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
      Index           =   7
      Left            =   7800
      TabIndex        =   35
      Top             =   4680
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bairro"
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
      Left            =   120
      TabIndex        =   34
      Top             =   4320
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Endereço"
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
      Left            =   120
      TabIndex        =   33
      Top             =   3960
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C..G.C./C.P.F.:"
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
      Left            =   4560
      TabIndex        =   32
      Top             =   3120
      Width           =   1470
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UF"
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
      Left            =   3240
      TabIndex        =   31
      Top             =   3120
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Placa"
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
      TabIndex        =   30
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
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
      Left            =   6840
      TabIndex        =   29
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transportadora "
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
      Left            =   120
      TabIndex        =   28
      Top             =   2640
      Width           =   1695
   End
End
Attribute VB_Name = "DadosOrcamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private LcNatureza As String
Private LcNota, LcBoleta, LcEspaco, LcLinha, LcEspC As String
Private LcSalto, LcQuant, a As Integer

Private Sub Command1_Click()
On Error Resume Next
FrmSelecioTransportadora.Show , Me

End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
If KeyCode = 113 Then SendKeys "%+{I}"
If KeyCode = 114 Then SendKeys "%+{D}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Command2_Click()
On Error Resume Next
Unload Me

End Sub
Function BuscaDados()
On Error Resume Next
Dim RsOrc As Recordset
Dim bb As Database

LcSql2 = "Select * from Orcamento where doc='" & orcamento.Documento.Text & "'"
Set bb = OpenDatabase(GLBase, False, False)

Set RsOrc = bb.OpenRecordset(LcSql2, dbOpenDynaset, dbSeeChanges, dbOptimistic)
PrevCommisao.Text = orcamento.emissao.Text
If Not RsOrc.EOF Then
   Txt(0).Text = RsOrc!Transp
   If RsOrc!TIPOTRANS = 1 Then
      Tipo.Text = "1- CIF"
   Else
      Tipo.Text = "2 - FOB"
   End If
   Placa.Text = RsOrc!PLACATRANS
   Txt(1).Text = RsOrc!UFTRANS
   Txt(2).Text = RsOrc!CGCCPFTRAN
   Txt(3).Text = RsOrc!ENDTRANS
   Txt(4).Text = RsOrc!MUNICTRANS
   Txt(6).Text = RsOrc!Cidade
   Txt(12).Text = RsOrc!Cep
   Txt(5).Text = RsOrc!UFMUNIC
   Txt(10).Text = RsOrc!FoneTransp
   Txt(7).Text = RsOrc!OBS02
   Txt(8).Text = RsOrc!OBS03
   Txt(9).Text = RsOrc!OBS04
   PrevCommisao.Text = RsOrc!PrevComissao
   TipoPag.Text = RsOrc!condpag
   TipoMonetario.Text = RsOrc!formapag
   Txt(11).Text = RsOrc!Dias
   Vencimento(0).Text = RsOrc!Vencimento1
   Vencimento(1).Text = RsOrc!vencimento2
   Vencimento(2).Text = RsOrc!vencimento3
   Vencimento(3).Text = RsOrc!vencimento4
   Vencimento(4).Text = RsOrc!vencimento5
   
End If
If TipoPag.Text = "A Vista" Then
   Vencimento(1).Visible = False
   Vencimento(2).Visible = False
   Vencimento(3).Visible = False
   Vencimento(4).Visible = False
Else
   Vencimento(1).Visible = True
   Vencimento(2).Visible = True
   Vencimento(3).Visible = True
   Vencimento(4).Visible = True
End If
GeraValor
RsOrc.Close
bb.Close
Set RsOrc = Nothing
Set bb = Nothing
End Function
Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
If KeyCode = 113 Then SendKeys "%+{I}"
If KeyCode = 114 Then SendKeys "%+{D}"
If KeyCode = 121 Then SendKeys "%+{F}"

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{I}"
If KeyCode = 121 Then SendKeys "%+{F}"
Visualizar.Visible = GLPadraoWindows
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim LcVer, a As Integer
Vencimento(0).Text = "  /  /  "
GlDadosTransportadora = False
If Not GlDadosTransportadora Then
  Me.Height = 2775
  For a = 0 To 20
    If a <> 11 Then Txt(a).Visible = False
     Label1(a).Visible = False
  Next
  
End If

Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
CarregaTipoMonetario
BuscaDados

End Sub

Function GeraValor() As Currency
On Error Resume Next
Dim LcValor As Currency

If Vencimento(4).Text = "  /  /  " Then
   If Vencimento(3).Text = "  /  /  " Then
      If Vencimento(2).Text = "  /  /  " Then
         If Vencimento(1).Text = "  /  /  " Then
            If Vencimento(0).Text = "  /  /  " Then
            Else
               valor.Text = CCur(orcamento.TotalOrcamento.Text)
               LcQuant = 1
               
            End If
         Else
            valor.Text = CCur(orcamento.TotalOrcamento.Text) / 2
            LcQuant = 2
         End If
      Else
         valor.Text = CCur(orcamento.TotalOrcamento.Text) / 3
         LcQuant = 3
      End If
   Else
      valor.Text = CCur(orcamento.TotalOrcamento.Text) / 4
      LcQuant = 4
   End If
Else
   valor.Text = CCur(orcamento.TotalOrcamento.Text) / 5
   LcQuant = 5
End If

Quantidade.Text = LcQuant
End Function
Function CarregaTipoMonetario()
On Error Resume Next
Dim RsMoney As Recordset
Dim bb As Database

TipoMonetario.Clear
Set bb = OpenDatabase(GLBase, False, False)

Set RsMoney = bb.OpenRecordset("Select * from alid008 where VENDA='S' order by XTPMONET", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Do Until RsMoney.EOF
   TipoMonetario.AddItem RsMoney("XTPMONET")
   RsMoney.MoveNext
Loop
TipoMonetario.AddItem "Nenhum"
RsMoney.Close
bb.Close
Set RsMoney = Nothing
Set bb = Nothing


End Function



Private Sub Imprime_Click()
LcCap = DadosOrcamento.Caption
DadosOrcamento.Caption = "Aguarde, Gerando Pedido."
orcamento.SalvaOrcamento
If Not GLPadraoWindows Then
  If Not GlMeiaFolha Then
     If orcamento.Natureza.Text <> "Orçamento" Then
      If GlOpcaoEmpresa = "olinto" Then
         orcamento.InprimeNotaOLinto
      Else
         orcamento.Imprimeorcamento
      End If
     Else
         orcamento.Imprimeorcamento
     End If
   Else
     LeConfiguracaoMeiaFolha
     For a = 1 To GlImpressaoMeiaFolha
        orcamento.ImprimeMeiaFolha
     Next
   End If
Else
  orcamento.ImprimeorcamentoWindows

End If


ConfirmaOrcamento.Show , Me
orcamento.limpanota
DadosOrcamento.Caption = LcCap


'FrmPedido.Txt(2).SetFocus
End Sub

Private Sub Imprime_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
If KeyCode = 113 Then SendKeys "%+{I}"
If KeyCode = 114 Then SendKeys "%+{D}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Placa_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
If KeyCode = 113 Then SendKeys "%+{I}"
If KeyCode = 114 Then SendKeys "%+{D}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub PrevCommisao_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
If KeyCode = 113 Then SendKeys "%+{I}"
If KeyCode = 114 Then SendKeys "%+{D}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub PrevCommisao_LostFocus()
If PrevCommisao.Text = "  /  /  " Then Exit Sub
If Not IsDate(PrevCommisao.Text) Then
   MsgBox "Digite Uma data Válida...", vbInformation, "Aviso"
   PrevCommisao.SetFocus
   Exit Sub
End If

End Sub

Private Sub Quantidade_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
If KeyCode = 113 Then SendKeys "%+{I}"
If KeyCode = 114 Then SendKeys "%+{D}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Tipo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
If KeyCode = 113 Then SendKeys "%+{I}"
If KeyCode = 114 Then SendKeys "%+{D}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub TipoMonetario_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
If KeyCode = 113 Then SendKeys "%+{I}"
If KeyCode = 114 Then SendKeys "%+{D}"
If KeyCode = 121 Then SendKeys "%+{F}"

End Sub

Private Sub TipoPag_Click()
On Error Resume Next
If TipoPag.Text = "A Vista" Then
   Vencimento(1).Visible = False
   Vencimento(2).Visible = False
   Vencimento(3).Visible = False
   Vencimento(4).Visible = False
Else
   Vencimento(1).Visible = True
   Vencimento(2).Visible = True
   Vencimento(3).Visible = True
   Vencimento(4).Visible = True
End If
'calculaprazo
End Sub

Private Sub TipoPag_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
If KeyCode = 113 Then SendKeys "%+{I}"
If KeyCode = 114 Then SendKeys "%+{D}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
If KeyCode = 113 Then SendKeys "%+{I}"
If KeyCode = 114 Then SendKeys "%+{D}"
If KeyCode = 121 Then SendKeys "%+{F}"""
End Sub

Private Sub Txt_LostFocus(Index As Integer)
On Error Resume Next
Txt(Index).Text = UCase(Txt(Index).Text)
If Index = 11 Then calculaprazo
End Sub
Function calculaprazo()
On Error Resume Next

Dim LCLEtra As String
Dim LcPrazo As String
Dim LcVencimento As Date
Dim LcValor As Currency
Dim LcControle, a As Integer
Dim LcTamanho As Long
LcControle = 0
LcTamanho = Len(Txt(11).Text)
For a = 1 To LcTamanho
   LCLEtra = Mid(Txt(11).Text, a, 1)
   If IsNumeric(LCLEtra) Then
      LcPrazo = LcPrazo & LCLEtra
   Else
      If LcControle = 4 Then
         MsgBox "Estao Disponíveis Somente Cinco Prazos de Pag.", 64, "AVISO"
         Txt(11).SetFocus
         Exit Function
      End If
      LcVencimento = Date + CLng(LcPrazo)
      Vencimento(LcControle).Text = Format(LcVencimento, "dd/mm/yy")
      LcControle = LcControle + 1
      valor.Text = CCur(orcamento.TotalOrcamento.Text) / LcControle
      LcQuant = LcControle
      Quantidade.Text = LcQuant
      LcPrazo = ""
   End If
Next
If Len(LcPrazo) > 0 Then
   LcVencimento = Date + CLng(LcPrazo)
   Vencimento(LcControle).Text = Format(LcVencimento, "dd/mm/yy")
   LcControle = LcControle + 1
   valor.Text = CCur(orcamento.TotalOrcamento.Text) / LcControle
   LcQuant = LcControle
   Quantidade.Text = LcQuant
   LcPrazo = ""
End If
End Function

Private Sub valor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
If KeyCode = 113 Then SendKeys "%+{I}"
If KeyCode = 114 Then SendKeys "%+{D}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Vencimento_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
If KeyCode = 113 Then SendKeys "%+{I}"
If KeyCode = 114 Then SendKeys "%+{D}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Vencimento_LostFocus(Index As Integer)
GeraValor
End Sub
