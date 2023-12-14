Attribute VB_Name = "Protecao"
Public Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTtoalNumberOfClusters As Long) As Long
Public Declare Function GetCurrentDirectory Lib "kernel32" Alias "GetCurrentDirectoryA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Declare Function getvolumeinformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpvolumenamebuffer As String, ByVal nvolumenamesize As Long, lpvolumeserialnumber As Long, lpmaximumcomponentlength As Long, lpfilesystemflags As Long, ByVal lpfilesystemnamebuffer As String, ByVal nfilesystemnamesize As Long) As Long

Public GLSerieHd, GlSerie, GlDirWindows, GlSerieSistema As String
Public GlSpacoHd As Long

Public Const NomeSys = "WinSystenl.ini"



Function SerieHd()
Dim lVSN As Long, n As Long, s1 As String, s2 As String
Dim unidad As String
Dim sTmp As String
On Error GoTo ErrorSerieHd

s1 = String$(255, Chr$(0))
s2 = String$(255, Chr$(0))
n = getvolumeinformation("c:\", s1, Len(s1), lVSN, 0, 0, s2, Len(s2))
sTmp = Hex$(lVSN)
Get_Number_serie = Left$(sTmp, 4) & "-" & Right$(sTmp, 4)
GLSerieHd = Get_Number_serie
Exit Function
ErrorSerieHd:
Resume Next
End Function
Function DirWindows()
On Error Resume Next

Dim LcDirWindows As String, LcCaracter, GlDirWinSystem As String
Dim retValue As Long, i As Integer
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
GlDirWindows = GlDirWinSystem

End Function
Function VerificaArquivo() As Integer
On Error Resume Next
Dim NomeArquivo As String

NomeArquivo = GlDirWindows & "\" & NomeSys

Open NomeArquivo For Input As #1

If err = 0 Then
   VerificaArquivo = True
Else
   VerificaArquivo = False
End If
Close #1
End Function

Function TamanhoHd()
On Error Resume Next

Dim LcSetor, LcLivre, LcTotal, LcBytes, LcRespostas, LcSpaco, LcSpacoLivre As Long
LcRespostas = GetDiskFreeSpace("C:\", LcSetor, LcBytes, LcLivre, LcTotal)
LcSpaco = (LcSetor * LcBytes * LcTotal) / 1000
LcSpacoLivre = (LcSetor * LcBytes * LcLivre) / 1000
GlSpacoHd = Int(LcSpaco)
End Function
Function CriaSystem() As Integer
Dim NomeArquivo As String
On Error GoTo ErrCria

NomeArquivo = GlDirWindows & "\" & NomeSys
MsgBox "Insira o Último Disket No Drive << A >> ", 48, "Aviso"
FileCopy "A:\SetupLoc.ini", NomeArquivo

Open "a:\Setuploc.ini" For Append As #1      ' Open file for output.
Write #1, Codifica("N")
Write #1, Codifica(GlSerieSistema)
Close #1

Open NomeArquivo For Append As #1      ' Open file for output.
Write #1, Codifica("S")
Write #1, Codifica(GlSerieSistema)
Write #1, Codifica(App.Path)
Write #1, Codifica(CStr(GLSerieHd))
Write #1, Codifica(CStr(GlSpacoHd))
Close #1

Exit Function
ErrCria:
LcResposta = MsgBox("Erro durante a Copia de Arquivos..", 21, "Erro Leitura")
If LcResposta = 4 Then Resume 0 Else End


End Function
Function Codifica(VarGeradan As String)
Dim Tamnho As Long, Tamanhoc As Long, vteg As Long
Dim Vg As String, VarGerada As String
Dim tec As Integer

Tamnho = Len(VarGeradan)

Tamanhoc = 1


Do While Tamanhoc <= Tamnho
    Vg = Mid(VarGeradan, Tamanhoc, 1)
    tec = Asc(Vg)
    vteg = tec * 2
    VarGerada = VarGerada & Chr(vteg)
    
    Tamanhoc = Tamanhoc + 1
Loop

Codifica = VarGerada

End Function
Function VerificaLocado() As Integer
On Error Resume Next
Dim RsProtecao As Recordset
Dim LcLocado As Integer

Set Dbbase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
Set RsProtecao = Dbbase.OpenRecordset("Seguranca", dbOpenTable, dbSeeChanges, dbOptimistic)
LcLocado = RsProtecao!Locado

RsProtecao.Close
Dbbase.Close

VerificaLocado = LcLocado

End Function

Function VerificaTravado() As Integer
On Error Resume Next
Dim RsProtecao As Recordset
Dim LcTravado As Integer

Set Dbbase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
Set RsProtecao = Dbbase.OpenRecordset("Seguranca", dbOpenTable, dbSeeChanges, dbOptimistic)
LcTravado = RsProtecao!Travado

RsProtecao.Close
Dbbase.Close

VerificaTravado = LcTravado


End Function
Function TravaSistema() As Integer
On Error Resume Next
Dim RsProtecao As Recordset
Dim LcTravado As Integer

Set Dbbase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
Set RsProtecao = Dbbase.OpenRecordset("Seguranca", dbOpenTable, dbSeeChanges, dbOptimistic)

RsProtecao.Edit
RsProtecao!Travado = True
RsProtecao.Update

RsProtecao.Close
Dbbase.Close
If err = 0 Then
  TravaSistema = True
Else
  TravaSistema = False
End If
  

End Function
Function DeveTravar() As Integer
On Error Resume Next
Dim RsProtecao As Recordset
Dim LcTravado As Integer

Set Dbbase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
Set RsProtecao = Dbbase.OpenRecordset("Seguranca", dbOpenTable, dbSeeChanges, dbOptimistic)
If RsProtecao!DataTravamento <= DateValue(GlDataSistema) Then
   DeveTravar = True
   RsProtecao.Edit
   RsProtecao!Travado = True
   RsProtecao.Update
Else
  DeveTravar = False
End If

RsProtecao.Close
Dbbase.Close
If err = 0 Then
  'DeveTravar = True
Else
 'DeveTravar = False
End If
End Function
Function VerificaProtecao()
On Error Resume Next

Dim RsProtecao As Recordset
Dim LcTravado As Integer

Set Dbbase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
Set RsProtecao = Dbbase.OpenRecordset("Seguranca", dbOpenTable, dbSeeChanges, dbOptimistic)

DirWindows
SerieHd
TamanhoHd
FrmPrincipal.Serie.Caption = "Com o Nº de Série " & GlSerieSistema
BuscaEmpresa
FrmPrincipal.Visible = True
RsProtecao.Close
Dbbase.Close
Exit Function
If Not IsNull(RsProtecao!SerieSistema) Then
   GlSerieSistema = RsProtecao!SerieSistema
Else
   MsgBox "Sistema Instalado Indevidamente...", 48, "Aviso"
   End
End If
If Not RsProtecao!Instalado Then
   Instala
Else
   protecao
End If
FrmPrincipal.Serie.Caption = "Com o Nº de Série " & GlSerieSistema
BuscaEmpresa
FrmPrincipal.Visible = True
RsProtecao.Close
Dbbase.Close
End Function
Function protecao()

If Not VerificaArquivo Then
   MsgBox "Sistema Instalado Indevidamente...", 48, "Aviso"
   End
End If

'If Not Verificahd Then
'   MsgBox "Sistema Instalado Indevidamente...", 48, "Aviso"
 '  End
'End If
If VerificaLocado Then
   
   DeveTravar
   If VerificaTravado Then
      MsgBox "Perido de Utilização Terminado." & Chr(13) & "Entre em contato com o Distribuidor...", 48, "Aviso"
      End
   End If
End If
End Function

Function Instala()
If LeArquivoTexto = "S" Then
   If Not TotalInstalacao Then
      GravaProtecao
      CriaSystem
      AcrescentaInstalacao
   Else
      MsgBox "Total de Instalação Já Ultrapassado...", 48, "Aviso"
      End
   End If
Else
   If Not TotalInstalacao Then
      GravaProtecao
      CriaSystem
      AcrescentaInstalacao
   Else
      MsgBox "Instalação do Sistema Indevida...", 48, "Aviso"
      End
   End If
End If
End Function
Function GravaProtecao()
On Error Resume Next
Dim RsProtecao As Recordset
Dim LcTravado As Integer

Set Dbbase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
Set RsProtecao = Dbbase.OpenRecordset("Seguranca", dbOpenTable, dbSeeChanges, dbOptimistic)
RsProtecao.Edit
RsProtecao!Instalado = True
RsProtecao.Update

End Function
Function Verificahd() As Integer
On Error Resume Next
Dim a, l, i, LcLocado As Integer
Dim LcTamnhoHd, LcSerie As String
Dim LcSpaco As Long
NomeArquivo = GlDirWindows & "\" & NomeSys
Open NomeArquivo For Input As #1      ' Open file for output.
a = 0
Stop
Do Until EOF(1)
   Input #1, LcCaracter    ' Read data into two variables.
   l = Len(LcCaracter)
   For i = 1 To l
      Caracter = Mid(LcCaracter, i, 1)
      Select Case a
        Case Is = 0
        Case Is = 1
        Case Is = 2
          tec = Asc(Caracter)
          LcSerie = LcSerie & Chr(tec / 2)
        Case Is = 3
          tec = Asc(Caracter)
          LcSerie = LcSerie & Chr(tec / 2)
        Case Is = 4
          tec = Asc(Caracter)
          LcTamnhoHd = LcTamnhoHd & Chr(tec / 2)
          LcSpaco = CLng(LcTamnhoHd)
      End Select
      
   Next
   Stop
   a = a + 1
Loop
Stop
Close #1



If LcSpaco = GlSpacoHd Then
      Verificahd = True
Else
     Verificahd = False
     Exit Function
End If

If LcSerie = GLSerieHd Then
   Verificahd = True
Else
   Verificahd = False
End If
      
End Function
Function BuscaEmpresa()
On Error Resume Next
Dim RsEmpresa As Recordset
Dim LcEmpresa As String
Set Dbbase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
Set RsEmpresa = Dbbase.OpenRecordset("Empresa", dbOpenDynaset, dbSeeChanges, dbOptimistic)
RsEmpresa.Edit
RsEmpresa!RazaoDaEmpresa = "World Video"
RsEmpresa.Update
LcEmpresa = "Licenciado Para a Empresa "
LcEmpresa = LcEmpresa & RsEmpresa!RazaoDaEmpresa
FrmPrincipal.LbEmpresa.Caption = LcEmpresa
RsEmpresa.Close

End Function
Function TotalInstalacao() As Integer
On Error Resume Next
Dim RsProtecao As Recordset
Dim LcTravado As Integer

Set Dbbase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
Set RsProtecao = Dbbase.OpenRecordset("Seguranca", dbOpenTable, dbSeeChanges, dbOptimistic)

If RsProtecao!TotalInstalcao > RsProtecao!Micros Then
   TotalInstalacao = True
Else
   TotalInstalacao = False
End If

End Function
Function AcrescentaInstalacao()
On Error Resume Next

Dim RsProtecao As Recordset
Dim LcTravado As Integer
Dim LcInstalacao As Long

Set Dbbase = OpenDatabase(GLBase, False, False, ";Pwd=muralha")
Set RsProtecao = Dbbase.OpenRecordset("Seguranca", dbOpenTable, dbSeeChanges, dbOptimistic)

If IsNull(RsProtecao!TotalInstalcao) Or RsProtecao!TotalInstalcao = 0 Then
   LcInstalacao = 1
Else
   LcInstalacao = RsProtecao!TotalInstalcao + 1
End If

RsProtecao.Edit
RsProtecao!TotalInstalcao = LcInstalacao
RsProtecao!SerieHd = GLSerieHd
RsProtecao!TamanhoHd = GlSpacoHd
RsProtecao.Update

RsProtecao.Close


End Function
Function LeArquivoTexto() As String
Dim LcCaracter, Caracter As String
Dim tec As Long
On Error GoTo ErrLerArqA
Open "a:\Setuploc.ini" For Input As #1      ' Open file for output.
Input #1, LcCaracter    ' Read data into two variables.

Caracter = Mid(LcCaracter, 1, 1)
tec = Asc(Caracter)
LeArquivoTexto = Chr(tec / 2)

Close #1
Exit Function
ErrLerArqA:
LcResposta = MsgBox("Insira o Último Diskete da Instalação no Drive <<A>>.", 33, "Termino da Instalação")
If LcResposta = 1 Then Resume 0 Else End
End Function
