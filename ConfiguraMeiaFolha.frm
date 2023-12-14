VERSION 5.00
Begin VB.Form ConfiguraMeiaFolha 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configura Impressão em meia Folha"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   4185
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Fechar"
      Height          =   495
      Left            =   1920
      TabIndex        =   10
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Salvar"
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CheckBox cabecalho 
      Caption         =   "Imprime Cabeçalho"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   3480
      Width           =   3735
   End
   Begin VB.TextBox impressao 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2760
      Width           =   3735
   End
   Begin VB.TextBox SaltoFinal 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   3735
   End
   Begin VB.TextBox margem 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   3735
   End
   Begin VB.TextBox itens 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   3735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Número de Impressão"
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
      Left            =   240
      TabIndex        =   8
      Top             =   2520
      Width           =   2310
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Linhas a Saltar no Final da Folha"
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
      TabIndex        =   7
      Top             =   1800
      Width           =   3420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Margem"
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
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quantidades de Itens Por Folha"
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
      Left            =   240
      TabIndex        =   5
      Top             =   360
      Width           =   3285
   End
End
Attribute VB_Name = "ConfiguraMeiaFolha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Dim Lcarq As Integer
Dim LcCaminho As String
Dim a As Integer

Lcarq = FreeFile

For a = Len(GLBase) To 1 Step -1
   If Mid(GLBase, a, 1) = "\" Then Exit For
Next
LcCaminho = Mid(GLBase, 1, a)
LcCaminho = LcCaminho & "meiaFolha.ini"
Open LcCaminho For Output As #Lcarq

If Len(itens.Text) = 0 Then itens.Text = 0
If Len(margem.Text) = 0 Then margem.Text = 0
If Len(SaltoFinal.Text) = 0 Then SaltoFinal.Text = 0
If Len(impressao.Text) = 0 Then impressao.Text = 0

Print #Lcarq, itens.Text '1
Print #Lcarq, margem.Text  '2
Print #Lcarq, SaltoFinal.Text '3
Print #Lcarq, impressao.Text '1
Print #Lcarq, cabecalho.Value  '3
Close #Lcarq

Unload Me

End Sub

Private Sub Command2_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
LeConfiguracaoMeiaFolha
itens.Text = GlItensMeiaFolha
margem.Text = GlMargemMeiaFolha
SaltoFinal.Text = GlSaltoFinalMeiaFolha
impressao.Text = GlImpressaoMeiaFolha

If GlCabecalhoMeiaFolha Then cabecalho.Value = 1 Else cabecalho.Value = 0


End Sub
