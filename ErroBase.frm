VERSION 5.00
Begin VB.Form ErroBase 
   BackColor       =   &H008080FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Banco de Dados Não Encontrado"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6780
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "&Recuperar do Backup"
      Height          =   615
      Left            =   4440
      TabIndex        =   8
      Top             =   3600
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Localizar o Banco de Dados"
      Height          =   615
      Left            =   2280
      TabIndex        =   7
      Top             =   3600
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Sair  do Sistema (Tentar Reconectar Depois)"
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "O que Você Deseja Fazer ?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   5
      Left            =   1680
      TabIndex        =   5
      Top             =   3000
      Width           =   3330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Se Você Não Utilize Sistema de Rede, Verifique"
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
      Index           =   4
      Left            =   600
      TabIndex        =   4
      Top             =   1800
      Width           =   5790
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caso não, Refaça a Conexão da Rede."
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
      Index           =   3
      Left            =   960
      TabIndex        =   3
      Top             =   1320
      Width           =   4695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "se o Diretório do Banco de Dados Foi Alterado."
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
      Index           =   2
      Left            =   360
      TabIndex        =   2
      Top             =   2370
      Width           =   5700
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Verifique se a Conexão de Rede está Funcionando "
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
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   864
      Width           =   6210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "O Banco de Dados Não Foi Encontrado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   4755
   End
End
Attribute VB_Name = "ErroBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=== PARA CAPTURAR A PASTA DESEJADA ========================================
'necessário para acionar o browser
Private Type tProcuraInformação
    hWndOwner As Long
    pidlRoot As Long
    sDisplayName As String
    sTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Private a As Integer
Private Declare Function SHBrowseForFolder Lib "Shell32.dll" (bBrowse As tProcuraInformação) As Long
Private Declare Function SHGetPathFromIDList Lib "Shell32.dll" (ByVal lL_Item As Long, ByVal sDir As String) As Long
'===========================================================================

'Private CAB As cCAB
Private Sub Command1_Click()
On Error Resume Next
End
End Sub

Private Sub Command2_Click()
On Error Resume Next
Unload Me
GlAchaBase = True
End Sub

Private Sub Command3_Click()
On Error Resume Next
Dim LcResposta      As Long
Dim LcCriterio      As String
Dim LCLEtra         As String
Dim LcArq           As String
Dim LcCaminho       As String
Dim LcAchou         As Boolean
Dim a               As Integer
Dim LcCamiMak       As String
Dim LcCamiCopia     As String
Dim sDiretorio      As String
Dim LcCap           As String
MsgBox "Para Efetuar a Restauração da Base de dados," & Chr(13) & "Todos as Máquinas deverão estar fora do Sistema.", vbInformation, "Aviso"
LcResposta = MsgBox("Esta Operação Irá Sobrescrever o Seu Banco de dados Atual," & Chr(13) & "As Alterações Feitas depois Deste Backup" & Chr(13) & Chr(13) & " SERÃO PERDIDAS." & Chr(13) & Chr(13) & "Confirma a Restauração ?", vbCritical + vbYesNo, "AVISO IMPORTANTE.")
If LcResposta = 7 Then
  MsgBox " Operação Cancelada pelo Usuário.", 64, "Aviso"
  Exit Sub
End If

'LcCamiMak = App.Path & "\bak" & Format(Date, "ddmmyy") & ".mdb"
'X = CopyFile(Trim$(GLBase), Trim(LcCamiMak), False)

LcCap = Me.Caption
Me.Caption = "Aguarde a finalização da Restauração dos dados...."
DoEvents
LcCamiMak = BuscaDirWin & "\Makecab.exe"
LcCamiCopia = App.Path & "\makecab.exe"
' ===Verifica se existe o Arquivo de compactacao no diretorio Windows
If Dir(LcCamiMak) = "" Then
   x = CopyFile(Trim$(LcCamiCopia), Trim(LcCamiMak), False)
   LcCamiMak = BuscaDirWin & "\Extract.exe"
   LcCamiCopia = App.Path & "\Extract.exe"
   x = CopyFile(Trim$(LcCamiCopia), Trim(LcCamiMak), False)
End If

'captura escolha de diretório pelo usuário...
'sDiretorio = sProcuraPorDiretório("Diretório para descompressão de arquivos")
For a = Len(GLBase) To 1 Step -1
    LCLEtra = Mid(GLBase, a, 1)
    If LCLEtra = "\" And Not LcAchou Then
       LcAchou = True
    Else
       If Not LcAchou Then
          LcArq = LCLEtra & LcArq
       Else
          LcCaminho = LCLEtra & LcCaminho
       End If
    End If
Next
sDiretorio = "C:\cabvir"
'se algum foi escolhido...
If sDiretorio <> "" Then

    'verifica contrabarra no caminho...
    sDiretorio = sFormataCaminho(sDiretorio)
    
    'chama rotuina de descompactação...
    Call Cab.Descomprimir(sDiretorio)
End If
GlRecuperou = True
Me.Caption = LcCap
Unload Me

GlAchaBase = True

End Sub

Private Sub Form_Load()
On Error Resume Next
'Set CAB = New cCAB
Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2
End Sub
