VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
arqini = BuscaDirWin
arqini = App.EXEName & ".ini"
GlNomeProjeto = App.EXEName & "RecLidis"
'If Dir(arqini, vbArchive) = "" Then
NomedoArquivo = arqini
abreconexao
Dim StrSql As String
Dim Rs As ADODB.Recordset
Dim RsContasReceber As ADODB.Recordset
StrSql = "Select * from alid050"
Set Rs = AbreRecordset(StrSql)

Do Until Rs.EOF
   StrSql = "Select * from alid015 where nf like '" & Rs!numNF & "%'"
   Set RsContasReceber = AbreRecordset(StrSql)
   If RsContasReceber.EOF Then
      RsContasReceber.AddNew
            
      RsContasReceber!NF = Rs!numNF & "/01"
      RsContasReceber!Cliente = Rs!Cliente
      RsContasReceber!TPMONET = "05"
      RsContasReceber!Valor = Rs!ValorNota
      RsContasReceber!Data = Rs!dtEmis
      RsContasReceber("DTVENC") = CDate(Rs!dtEmis) + 30
                
                'RsContasReceber!DTPAGTO = Format(txt(12).Text, "dd/mm/yy")
      RsContasReceber!VALPAGO = 0
      RsContasReceber!tipord = "R"
      RsContasReceber!acrescimo = 0
      RsContasReceber.Update
   
   End If

   Rs.MoveNext
Loop

End Sub
Function BuscaDirWin() As String
Dim LcDirWindows    As String
Dim LcCaracter      As String
Dim GlDirWinSystem  As String
Dim retValue        As Long
Dim i               As Integer
Dim GlBuffer        As Integer
Dim GlDevApi        As Integer
GlBuffer = 255
LcDirWindows = String(255, " ")

GlDevApi = GetWindowsDirectory(LcDirWindows, GlBuffer)

'=== Esta Sequencia Tem por Finalidade, Separar somente o
'=== Nome do Diretorio Windows, Separando os Caracters Indesejados
'=== Os Caracteres ascII de 47 a 126 são as Letras Validas

For i = 1 To 255
       LcCaracter = Mid(LcDirWindows, i, 1)
       If Asc(LcCaracter) >= 47 And Asc(LcCaracter) <= 126 Then
          GlDirWinSystem = GlDirWinSystem & LcCaracter
       Else
          Exit For
       End If
Next i
BuscaDirWin = GlDirWinSystem
End Function

