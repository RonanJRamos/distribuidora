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
'�ltima modifica��o ===> Data: 12/11/2000 Hora: 10:01:53 <===
'Criado por: Gesiel Ferreira de Souza em 11/11/2000 �s 15:12:36 horas.
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
'necess�rio para acionar o browser
Private Type tProcuraInforma��o
    hWndOwner As Long
    pidlRoot As Long
    sDisplayName As String
    sTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type
Private Declare Function SHBrowseForFolder Lib "Shell32.dll" (bBrowse As tProcuraInforma��o) As Long
Private Declare Function SHGetPathFromIDList Lib "Shell32.dll" (ByVal lL_Item As Long, ByVal sDir As String) As Long
'===========================================================================

'Dim CAB As cCAB


Public Function sProcuraPorDiret�rio(sTitulo As String) As String
On Error Resume Next
'ACIONA O BROWSER A PROCURA DE DIRET�RIO

Dim oProcuraInforma��o      As tProcuraInforma��o
Dim lItem                   As Long
Dim sNomeDiret�rio          As String
   
oProcuraInforma��o.hWndOwner = hWnd
oProcuraInforma��o.pidlRoot = 0
oProcuraInforma��o.sDisplayName = Space$(260)
oProcuraInforma��o.sTitle = sTitulo
oProcuraInforma��o.ulFlags = 1 ' Retorna nome do diretorio.
oProcuraInforma��o.lpfn = 0
oProcuraInforma��o.lParam = 0
oProcuraInforma��o.iImage = 0

lItem = SHBrowseForFolder(oProcuraInforma��o)
If lItem Then
    sNomeDiret�rio = Space$(260)
    If SHGetPathFromIDList(lItem, sNomeDiret�rio) Then
        sProcuraPorDiret�rio = Left(sNomeDiret�rio, InStr(sNomeDiret�rio, Chr$(0)) - 1)
    Else
        sProcuraPorDiret�rio = ""
    End If
End If
End Function



Private Sub cmdComprimir_Click()
'MONTA LISTA DE ARQUIVOS A COMPRIMIR E CHAMA A ROTINA DE COMPRESS�O

Dim iNum            As Integer
Dim iLinha          As Integer
Dim sArquivo()      As String

'se existem arquivos selecionados...
If lstArquivos.SelCount > 0 Then
    'faz um loop pela lista...
    For iNum = 0 To lstArquivos.ListCount - 1
        'se este �tem est� selecionado...
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
    'ajusta as propriedades, se necess�rios
    
    'se o drive padr�o � A:\ , n�o precisa ser informado porque � o padr�o
    'CAB.BackupDrive = B
    
    'esta propriedade usa por padrao o diret�rio temp do Windows
    'mais se quiser, pode indicar outro...
    'CAB.PastaTemp = "C:\Temp"
    '=====================================
    
    'chama rotina de compacta��o com os par�metros escolhidos...
   ' Call CAB.Comprimir(cboBackup.Text = "Sim", lstArquivos.Tag, sArquivo())
Else
    MsgBox "Escolha arquivos os que deseja compactar", vbOKOnly + vbCritical, "Erro"
End If
End Sub


Private Sub cmdDescomprimir_Click()
'CAPTURA A PASTA E LISTA ARQUIVOS

'declara vari�vel
Dim sDiretorio      As String

'captura escolha de diret�rio pelo usu�rio...
sDiretorio = sProcuraPorDiret�rio("Diret�rio para descompress�o de arquivos")

'se algum foi escolhido...
If sDiretorio <> "" Then

    'verifica contrabarra no caminho...
    sDiretorio = sFormataCaminho(sDiretorio)
    
    'chama rotuina de descompacta��o...
    'Call CAB.Descomprimir(sDiretorio)
End If
 
End Sub


Private Sub CmdFechar_Click()
Unload Me
End Sub



Private Sub cmdPasta_Click()
'CAPTURA A PASTA E LISTA ARQUIVOS

'declara vari�vel
Dim sDiretorio      As String
lstArquivos.Clear
'captura escolha de diret�rio pelo usu�rio...
sDiretorio = sProcuraPorDiret�rio("Diret�rio dos arquivos que deseja comprimir")

'se algum foi escolhido...
If sDiretorio <> "" Then
    'lista arquivos...
    Call Proc_ListaArquivos(sDiretorio)
    lstArquivos.Tag = sDiretorio
End If
End Sub

Public Sub Proc_ListaArquivos(ByVal sPasta As String)
'FAZ LISTA DE ARQUIVOS DO DIRET�RIO ESCOLHIDO


'declara vari�vel
Dim sArquivo As String
sPasta = sFormataCaminho(sPasta)
'pega a primeira entrada...
sArquivo = Dir(sPasta, vbArchive)

'come�a o Loop enquanto nomes forem encontrados...
Do While sArquivo <> ""
    ' Ignora o diret�rio...
    If sArquivo <> "." And sArquivo <> ".." Then
        'verifica se � um arquivo
        If (GetAttr(sPasta & sArquivo) And vbArchive) = vbArchive Then
            'acrescenta este arquivo � lista...
            lstArquivos.AddItem sArquivo
        End If
    End If
    'captura o pr�ximo nome...
    sArquivo = Dir
Loop

End Sub

Private Sub Form_Load()
'cria nova ocorr�ncia da classe...
'Set CAB = New cCAB

'preenche controle
cboBackup.AddItem "Sim"
cboBackup.AddItem "N�o"
cboBackup.ListIndex = 0
End Sub


Private Sub Form_Unload(Cancel As Integer)
'If Not CAB Is Nothing Then
'    Set CAB = Nothing
'End Sub


