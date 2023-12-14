VERSION 5.00
Begin VB.Form frmCab 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CAB"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   4665
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstArquivos 
      Height          =   2985
      Left            =   360
      Style           =   1  'Checkbox
      TabIndex        =   7
      Top             =   360
      Width           =   3975
   End
   Begin VB.Frame fraCAB 
      Height          =   915
      Index           =   1
      Left            =   180
      TabIndex        =   5
      Top             =   4260
      Width           =   4335
      Begin VB.CommandButton cmdDescomprimir 
         Caption         =   "Descomprimir"
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   300
         Width           =   1215
      End
   End
   Begin VB.Frame fraCAB 
      Height          =   4155
      Index           =   0
      Left            =   180
      TabIndex        =   1
      Top             =   60
      Width           =   4335
      Begin VB.CommandButton cmdPasta 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3660
         TabIndex        =   8
         Top             =   3360
         Width           =   435
      End
      Begin VB.CommandButton cmdComprimir 
         Caption         =   "Comprimir"
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   3540
         Width           =   1215
      End
      Begin VB.ComboBox cboBackup 
         Height          =   315
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Fazer Backup"
         Height          =   195
         Left            =   1620
         TabIndex        =   4
         Top             =   3480
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "Fechar"
      Height          =   495
      Left            =   1860
      TabIndex        =   0
      Top             =   5340
      Width           =   1215
   End
End
Attribute VB_Name = "frmCab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Última modificação ===> Data: 12/11/2000 Hora: 10:01:53 <===
'Criado por: Gesiel Ferreira de Souza em 11/11/2000 às 15:12:36 horas.
'frmCab (Code)CAB
'*********************************************************************

'ESTE PROGRAMA DESTINA-SE A COMPACTAR E DESCOMPCTAR
'LISTAS DE ARQUIVOS EM VOLUMES CAB DE 1.44 MB CADA,
'UTILIZANDO OS PROGRAMAS MAKECAB.EXE E EXTRACT.EXE.

'IMPORTANTE: OS PROGRAMAS MAKECAB.EXE E EXTRACT.EXE
'DEVERAO SER COPIADOS PARA A PASTA DESTE PROJETO
'PARA QUE O MESMO POSSA FUNCIONAR

Option Explicit

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
Private Declare Function SHBrowseForFolder Lib "Shell32.dll" (bBrowse As tProcuraInformação) As Long
Private Declare Function SHGetPathFromIDList Lib "Shell32.dll" (ByVal lL_Item As Long, ByVal sDir As String) As Long
'===========================================================================

'Dim CAB As cCAB


Public Function sProcuraPorDiretório(sTitulo As String) As String
On Error Resume Next
'ACIONA O BROWSER A PROCURA DE DIRETÓRIO

Dim oProcuraInformação      As tProcuraInformação
Dim lItem                   As Long
Dim sNomeDiretório          As String
   
oProcuraInformação.hWndOwner = hWnd
oProcuraInformação.pidlRoot = 0
oProcuraInformação.sDisplayName = Space$(260)
oProcuraInformação.sTitle = sTitulo
oProcuraInformação.ulFlags = 1 ' Retorna nome do diretorio.
oProcuraInformação.lpfn = 0
oProcuraInformação.lParam = 0
oProcuraInformação.iImage = 0

lItem = SHBrowseForFolder(oProcuraInformação)
If lItem Then
    sNomeDiretório = Space$(260)
    If SHGetPathFromIDList(lItem, sNomeDiretório) Then
        sProcuraPorDiretório = Left(sNomeDiretório, InStr(sNomeDiretório, Chr$(0)) - 1)
    Else
        sProcuraPorDiretório = ""
    End If
End If
End Function



Private Sub cmdComprimir_Click()
'MONTA LISTA DE ARQUIVOS A COMPRIMIR E CHAMA A ROTINA DE COMPRESSÃO

Dim iNum            As Integer
Dim iLinha          As Integer
Dim sArquivo()      As String

'se existem arquivos selecionados...
If lstArquivos.SelCount > 0 Then
    'faz um loop pela lista...
    For iNum = 0 To lstArquivos.ListCount - 1
        'se este ítem está selecionado...
        If lstArquivos.Selected(iNum) = True Then
            'redimenciona o array...
            ReDim Preserve sArquivo(iLinha)
            'atualiza o valor deste membro...
            sArquivo(iLinha) = lstArquivos.List(iNum)
            'incrementa contador...
            iLinha = iLinha + 1
        End If
    Next
    
    '======================================
    'ajusta as propriedades, se necessários
    
    'se o drive padrão é A:\ , não precisa ser informado porque é o padrão
    'CAB.BackupDrive = B
    
    'esta propriedade usa por padrao o diretório temp do Windows
    'mais se quiser, pode indicar outro...
    'CAB.PastaTemp = "C:\Temp"
    '=====================================
    
    'chama rotina de compactação com os parâmetros escolhidos...
   ' Call CAB.Comprimir(cboBackup.Text = "Sim", lstArquivos.Tag, sArquivo())
Else
    MsgBox "Escolha arquivos os que deseja compactar", vbOKOnly + vbCritical, "Erro"
End If
End Sub


Private Sub cmdDescomprimir_Click()
'CAPTURA A PASTA E LISTA ARQUIVOS

'declara variável
Dim sDiretorio      As String

'captura escolha de diretório pelo usuário...
sDiretorio = sProcuraPorDiretório("Diretório para descompressão de arquivos")

'se algum foi escolhido...
If sDiretorio <> "" Then

    'verifica contrabarra no caminho...
    sDiretorio = sFormataCaminho(sDiretorio)
    
    'chama rotuina de descompactação...
    'Call CAB.Descomprimir(sDiretorio)
End If
 
End Sub


Private Sub CmdFechar_Click()
Unload Me
End Sub



Private Sub cmdPasta_Click()
'CAPTURA A PASTA E LISTA ARQUIVOS

'declara variável
Dim sDiretorio      As String
lstArquivos.Clear
'captura escolha de diretório pelo usuário...
sDiretorio = sProcuraPorDiretório("Diretório dos arquivos que deseja comprimir")

'se algum foi escolhido...
If sDiretorio <> "" Then
    'lista arquivos...
    Call Proc_ListaArquivos(sDiretorio)
    lstArquivos.Tag = sDiretorio
End If
End Sub

Public Sub Proc_ListaArquivos(ByVal sPasta As String)
'FAZ LISTA DE ARQUIVOS DO DIRETÓRIO ESCOLHIDO


'declara variável
Dim sArquivo As String
sPasta = sFormataCaminho(sPasta)
'pega a primeira entrada...
sArquivo = Dir(sPasta, vbArchive)

'começa o Loop enquanto nomes forem encontrados...
Do While sArquivo <> ""
    ' Ignora o diretório...
    If sArquivo <> "." And sArquivo <> ".." Then
        'verifica se é um arquivo
        If (GetAttr(sPasta & sArquivo) And vbArchive) = vbArchive Then
            'acrescenta este arquivo à lista...
            lstArquivos.AddItem sArquivo
        End If
    End If
    'captura o próximo nome...
    sArquivo = Dir
Loop

End Sub

Private Sub Form_Load()
'cria nova ocorrência da classe...
'Set CAB = New cCAB

'preenche controle
cboBackup.AddItem "Sim"
cboBackup.AddItem "Não"
cboBackup.ListIndex = 0
End Sub


Private Sub Form_Unload(Cancel As Integer)
'If Not CAB Is Nothing Then
'    Set CAB = Nothing
'End Sub


