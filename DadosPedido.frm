VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form DadosPedido 
   BackColor       =   &H00CAE1A2&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dados Finais do Pedido"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9360
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   9360
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   10
      Left            =   1920
      MaxLength       =   30
      TabIndex        =   5
      Top             =   1320
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar F10"
      Height          =   375
      Left            =   7920
      TabIndex        =   15
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton Imprime 
      Caption         =   "&Imprimir F2"
      Height          =   375
      Left            =   7920
      TabIndex        =   14
      Top             =   3000
      Width           =   1335
   End
   Begin VB.ComboBox TipoMonetario 
      Height          =   315
      Left            =   1920
      TabIndex        =   13
      Top             =   4320
      Width           =   3255
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   9
      Left            =   1920
      MaxLength       =   30
      TabIndex        =   12
      Top             =   3930
      Width           =   5655
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   8
      Left            =   1920
      MaxLength       =   30
      TabIndex        =   11
      Top             =   3645
      Width           =   5655
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   7
      Left            =   1920
      MaxLength       =   30
      TabIndex        =   10
      Top             =   3360
      Width           =   5655
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   6
      Left            =   1920
      MaxLength       =   12
      TabIndex        =   9
      Top             =   2520
      Width           =   5655
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   5
      Left            =   8280
      MaxLength       =   2
      TabIndex        =   8
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   4
      Left            =   1920
      MaxLength       =   20
      TabIndex        =   7
      Top             =   2160
      Width           =   5655
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   3
      Left            =   1920
      MaxLength       =   30
      TabIndex        =   6
      Top             =   1680
      Width           =   7095
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   2
      Left            =   6120
      MaxLength       =   18
      TabIndex        =   4
      Top             =   960
      Width           =   2895
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   1
      Left            =   3720
      MaxLength       =   2
      TabIndex        =   3
      Top             =   960
      Width           =   735
   End
   Begin MSMask.MaskEdBox Placa 
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   840
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
      ItemData        =   "DadosPedido.frx":0000
      Left            =   7440
      List            =   "DadosPedido.frx":000A
      TabIndex        =   1
      Text            =   "1- CIF"
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   0
      Left            =   1920
      MaxLength       =   30
      TabIndex        =   0
      Top             =   360
      Width           =   4815
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
      TabIndex        =   27
      Top             =   1320
      Width           =   540
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   7800
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line2 
      X1              =   7800
      X2              =   9360
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line1 
      X1              =   7800
      X2              =   7800
      Y1              =   2760
      Y2              =   4680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Pagamento"
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
      Left            =   120
      TabIndex        =   26
      Top             =   4320
      Width           =   1740
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
      TabIndex        =   25
      Top             =   2880
      Width           =   4185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inscr. Est."
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
      TabIndex        =   24
      Top             =   2520
      Width           =   1020
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
      TabIndex        =   23
      Top             =   2040
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Município"
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
      TabIndex        =   22
      Top             =   2160
      Width           =   1005
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
      TabIndex        =   21
      Top             =   1680
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
      TabIndex        =   20
      Top             =   960
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
      TabIndex        =   19
      Top             =   960
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
      Left            =   240
      TabIndex        =   18
      Top             =   960
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
      TabIndex        =   17
      Top             =   360
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
      TabIndex        =   16
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "DadosPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private LcNatureza As String
Private LcNota, LcBoleta, LcEspaco, LcLinha, LcEspC As String
Private LcSalto, LcQuant, a As Integer

Private Sub Command2_Click()
On Error Resume Next
Unload Me

End Sub

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{I}"
If KeyCode = 121 Then SendKeys "%+{F}"

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then SendKeys "%+{I}"
If KeyCode = 121 Then SendKeys "%+{F}"

End Sub

Private Sub Form_Load()
Dim LcVer As Integer

Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
CarregaTipoMonetario
End Sub

Function CarregaTipoMonetario()
Dim RsMoney As Recordset
TipoMonetario.Clear
AbreBase
Set RsMoney = Dbbase.OpenRecordset("Select * from alid008  order by XTPMONET")
Do Until RsMoney.EOF
   TipoMonetario.AddItem RsMoney("XTPMONET")
   RsMoney.MoveNext
Loop
RsMoney.Close
Dbbase.Close
Set RsMoney = Nothing
Set Dbbase = Nothing


End Function



Private Sub Imprime_Click()
FrmPedido.SalvaPedido
FrmPedido.imprimepedido
FrmPedido.limpanota
End Sub

Private Sub Imprime_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{I}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Placa_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{I}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Tipo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{I}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub TipoMonetario_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{I}"
If KeyCode = 121 Then SendKeys "%+{F}"

End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 113 Then SendKeys "%+{I}"
If KeyCode = 121 Then SendKeys "%+{F}"
End Sub

Private Sub Txt_LostFocus(Index As Integer)
On Error Resume Next
Txt(Index).Text = UCase(Txt(Index).Text)
End Sub




