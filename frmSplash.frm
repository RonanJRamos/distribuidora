VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00E7DAAD&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5190
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   8040
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E7DAAD&
      Height          =   5115
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7800
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E7DAAD&
         Height          =   1035
         Left            =   120
         Picture         =   "frmSplash.frx":0000
         ScaleHeight     =   975
         ScaleWidth      =   3015
         TabIndex        =   9
         Top             =   360
         Width           =   3075
      End
      Begin VB.Label Aviso 
         Alignment       =   2  'Center
         BackColor       =   &H00F4C280&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   4320
         Width           =   7095
      End
      Begin VB.Label Serie 
         AutoSize        =   -1  'True
         BackColor       =   &H00F4C280&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   3360
         TabIndex        =   12
         Top             =   1680
         Width           =   90
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00F4C280&
         BackStyle       =   0  'Transparent
         Caption         =   "TeleFax: (31) 3032-4980  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1440
         TabIndex        =   11
         Top             =   2640
         Width           =   5415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00F4C280&
         BackStyle       =   0  'Transparent
         Caption         =   "Aguarde, Sistema Carregando ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   2160
         TabIndex        =   10
         Top             =   3480
         Width           =   4095
      End
      Begin VB.Label lblCopyright 
         BackColor       =   &H00F4C280&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8,25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1680
         Width           =   2895
      End
      Begin VB.Label lblCompany 
         BackColor       =   &H00F4C280&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8,25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1920
         Width           =   2415
      End
      Begin VB.Label lblWarning 
         BackColor       =   &H00F4C280&
         BackStyle       =   0  'Transparent
         Caption         =   "Aviso: A Copia NÃO AUTORIZADA deste Produto Resultará em SANSÕES Previstas na Lei."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8,25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   360
         TabIndex        =   2
         Top             =   3960
         Width           =   7335
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00F4C280&
         BackStyle       =   0  'Transparent
         Caption         =   "Versão 4.60"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   3240
         TabIndex        =   5
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00F4C280&
         BackStyle       =   0  'Transparent
         Caption         =   "Decisão Tecnologia em Informática"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2160
         TabIndex        =   6
         Top             =   2340
         Width           =   4005
      End
      Begin VB.Label lblProductName 
         BackColor       =   &H00F4C280&
         BackStyle       =   0  'Transparent
         Caption         =   "Desenvolvido"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8,25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1050
         Left            =   3360
         TabIndex        =   8
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label Aviso1 
         Alignment       =   2  'Center
         BackColor       =   &H00F4C280&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8,25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   4800
         Width           =   6855
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2160
         TabIndex        =   7
         Top             =   480
         Width           =   105
      End
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   7095
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private a As Integer
Option Explicit
Private LcVeri As Integer
Private LcVerifreg As Boolean
Function ReparaBase()
Dim LcMsg As String

On Error Resume Next
'GLBase.Close
'Me.Caption = "Aguarde, Reparando a Base de dados..."
If Not Dbbase Is Nothing Then
   Set Dbbase = Nothing
End If
DBEngine.RepairDatabase GLBase
If Not Dbbase Is Nothing Then
   Set Dbbase = Nothing
End If
'MsgBox "Operação Terminada com Sucesso..."
End Function

Private Sub Form_Activate()
On Error Resume Next
Me.Refresh
NomeMaquina
'Desabilitatodos
'ReparaBase
'Load FrmPesquisaProdutos
'Load FrmPesquisaCliente
'Load FrmPesquisaFornecedores
Load FrmBuscaCliente
Load FrmBuscaProduto
'Load FichaEstoque
abreconexao
Unload frmSplash
If LcVeri Then
  frmLogin.Show
Else
  If Not LcVerifreg Then FrmAviso.Show
End If
'If GlServidorImpressora Then LogNotaFiscal
If GlServidorImpressoraOrc Then logOrcamento
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
 On Error Resume Next
  
  If Len(Trim(GlFormInicial)) = 0 Then
     VerificaOpcoes
  End If
  lblVersion.Caption = "Versão " & App.Major & "." & App.Minor & "." & App.Revision
  lblProductName.Caption = App.Title
  lblPlatform.Caption = "Decisão Tec. em Informática"
  LcVeri = VerificaProtecao
End Sub
Function VerificaProtecao() As Integer

VerificaProtecao = True
End Function

Private Sub Frame1_Click()
    Unload Me
End Sub

