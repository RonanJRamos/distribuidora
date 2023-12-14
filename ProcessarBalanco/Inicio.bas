Attribute VB_Name = "Inicio"
Declare Function GetPrivateProfileStringSections& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName&, ByVal lpszKey&, ByVal lpszDefault$, ByVal lpszReturnBuffer$, ByVal cchReturnBuffer&, ByVal lpszFile$)
Declare Function GetPrivateProfileKeys% Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal Section$, ByVal Zero&, ByVal Default$, ByVal ReturnBuffer$, ByVal LenReturnBuffer%, ByVal FileName$)
Declare Function GetPrivateProfileStringA Lib "kernel32" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileDelKey% Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal Section$, ByVal Entry As Any, ByVal Zero&, ByVal FileName$)
Declare Function WritePrivateProfileDelSect% Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal Section$, ByVal Zero&, ByVal EmptyStr$, ByVal FileName$)
Declare Function WritePrivateProfileStringA% Lib "kernel32" (ByVal Section$, ByVal Entry As Any, ByVal CharStr As Any, ByVal FileName$)


Public Declare Function Extenso Lib "Extens32.dll" Alias "extenso" (ByVal valor As String, ByVal Retorno As String) As Integer
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function ConsisteInscricaoEstadual Lib "DllInscE32.dll" (ByVal Insc As String, ByVal UF As String) As Integer

Public Arqini As String
Public GLBase As String
Public GlNomeMaquina As String
Public GlUsuario As String

Sub Main()
'GLBase = "\\Servidor\c\BancoDados\lidis.mdb"
GLBase = "D:\PROJETO\Lidis Sql\bancodedados\estoque fiscal lidis.mdb"
Arqini = BuscaDirWin
Arqini = Arqini & "\" & App.EXEName & ".ini"
abreconexao
ProcessaBalanco.Show
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
Function LeIni(Seção$, Chave$, Arqini$) As String
Dim i&, Ret$, max_size As Byte
  max_size = 255
  Ret = Space(max_size)
  i = GetPrivateProfileStringA(Seção$, Chave$, "", Ret, max_size, Arqini$)
  LeIni = Left(Ret, InStr(Ret, Chr(0)) - 1)
End Function
'Write in Ini File =============================================
Function GravaIni(Seção$, Chave$, strValue$, Arqini$) As Boolean
On Error GoTo errGrvini
Dim i&
  i = WritePrivateProfileStringA(Seção$, Chave$, strValue$, Arqini$)
  GravaIni = Len(LeIni(Seção, Chave$, Arqini$)) > 0
Exit Function
errGrvini:
GravaIni = False
End Function
'Delete Keys in Ini File ==========================
Function DelKey(Section$, Key$, Arq$) As Boolean
On Error GoTo ErrDelKey
  WritePrivateProfileDelKey Section, Key, 0&, Arq
  DelKey = True
Exit Function
ErrDelKey:
  DelKey = False
End Function
'Delete Sections in Ini File =================
Function DelSection(Section$, Arq$) As Boolean
On Error GoTo ErrDelKey
  WritePrivateProfileDelSect Section, 0&, "", Arq
  DelSection = True
Exit Function
ErrDelKey:
  DelSection = False
End Function
'Lê os nomes de todas as chaves de uma dada seção========
Function KeysInSection(Section$, Arq$) As String
  Dim Buff As String * 1024, Result%
  Result = GetPrivateProfileKeys(Section, 0&, "", Buff, Len(Buff), Arq)
  KeysInSection = Left$(Buff, Result)
End Function
'Lê os nomes de todas as seções de um dado arquivo========
Function SectionsInFile(Arq$) As String
  Dim Rtn$, Result$, Pos&
  Result = Chr(255)
  Rtn = Space(1024)
  success = GetPrivateProfileStringSections(0, 0, "", Rtn, 1024, Arq$)
  Pos = InStr(1, Rtn, "  ")
  SectionsInFile = Left$(Rtn, (Pos - 2)) 'Result)
End Function

