Attribute VB_Name = "Atualizabase"
Public LcIndices As Index
Sub ProcessaDDl()
Dim LcArquivo As String
Dim LocalArquivo As String
Dim NumeroArq As Integer
Dim linha As String
Dim StrSql As String
Dim Str() As String
Dim a As Integer
Dim NomeTb As String
Dim db As Database
On Error GoTo ErroProcessaDll
NumeroArq = FreeFile

LocalArquivo = App.Path & "\Tabelas.SQL"

If Dir(LocalArquivo, vbArchive) = "" Then Exit Sub

Open LocalArquivo For Input As #NumeroArq

Do Until EOF(NumeroArq)
   Line Input #NumeroArq, linha
   StrSql = StrSql & linha
Loop
Close #NumeroArq

StrSql = Replace(StrSql, Chr(13), "")
StrSql = Replace(StrSql, Chr(10), "")

'MsgBox StrSql
Str = Split(StrSql, ";")
Set db = OpenDatabase(GLBase)
ExibeMsgAtualizacao.Show
For a = 0 To UBound(Str)
   DoEvents
   StrSql = Str(a) & ";"
   NomeTb = Mid(StrSql, 1, InStr(1, StrSql, "("))
   NomeTb = Replace(Replace(Replace(Replace(NomeTb, "CREATE", ""), "TABLE", ""), " ", ""), "(", "")
  ' If InStr(UCase(NomeTb), "OS") > 0 Then Stop
   ExibeMsgAtualizacao.Caption = "Criando Tabela " & NomeTb
   Debug.Print StrSql
  StrSql = Replace(StrSql, Chr(13), "")
  StrSql = Replace(StrSql, Chr(10), "")

   db.Execute StrSql

Next
If Dir(LocalArquivo, vbArchive) <> "" Then Kill LocalArquivo
Unload ExibeMsgAtualizacao
Exit Sub
ErroProcessaDll:
'MsgBox err.Description & err.Number
 Resume Next
If err.Number = 3010 Then Resume Next

Debug.Print err.Description & err.Number
'3624-2536
End Sub
Sub ProcessaDDlMySql()
Dim LcArquivo As String
Dim LocalArquivo As String
Dim NumeroArq As Integer
Dim linha As String
Dim StrSql As String
Dim Str() As String
Dim a As Integer
Dim NomeTb As String
On Error GoTo ErroProcessaDll
NumeroArq = FreeFile

LocalArquivo = App.Path & "\TabelasMySql.SQL"

If Dir(LocalArquivo, vbArchive) = "" Then Exit Sub

Open LocalArquivo For Input As #NumeroArq

Do Until EOF(NumeroArq)
   Line Input #NumeroArq, linha
   StrSql = StrSql & linha
Loop
Close #NumeroArq

StrSql = Replace(StrSql, Chr(13), "")
StrSql = Replace(StrSql, Chr(10), "")

'MsgBox StrSql
Str = Split(StrSql, ";")
abreconexao
ExibeMsgAtualizacao.Show
For a = 0 To UBound(Str)
   DoEvents
   StrSql = Str(a) & ";"
   NomeTb = Mid(StrSql, 1, InStr(1, StrSql, "("))
   NomeTb = Replace(Replace(Replace(Replace(NomeTb, "CREATE", ""), "TABLE", ""), " ", ""), "(", "")
  ' If InStr(UCase(NomeTb), "OS") > 0 Then Stop
  ' Debug.Print StrSql
  StrSql = Replace(StrSql, Chr(13), "")
  StrSql = Replace(StrSql, Chr(10), "")

   ExecutaSql StrSql

Next
If Dir(LocalArquivo, vbArchive) <> "" Then Kill LocalArquivo
Unload ExibeMsgAtualizacao
Exit Sub
ErroProcessaDll:
'MsgBox Err.Description & Err.Number
 Resume Next
If err.Number = 3010 Then Resume Next

Debug.Print err.Description & err.Number
'3624-2536
End Sub
Public Function VerificaVersao()
On Error Resume Next

Dim LcVersao    As Long
Dim LcNumero    As Integer
Dim a           As Integer
Dim LcArq       As String
Dim LCLEtra     As String
Dim LcDir       As String
Dim LcVerTxt    As String
Dim LcNomeExecutavel As String
LcNomeExecutavel = App.EXEName

LcNumero = FreeFile
'=== Verifica o numero da Versao Atual
LcVersao = CLng(App.Major) + CLng(App.Minor) + CLng(App.Revision)
'=== Busca o local da Base de dados
For a = Len(GLBase) To 1 Step -1
    If Mid(GLBase, a, 1) = "\" Then
       Exit For
    End If
Next
'=== Monta o nome do Arquivo
LcDir = Mid(GLBase, 1, a)
LcArq = LcDir & "versao.dvc"
'=== Verifica se o arquivo existe
If Dir(LcArq) <> "" Then
   Open LcArq For Input As #LcNumero
   Line Input #LcNumero, LcVerTxt
   Close #LcNumero
   If CLng(LcVerTxt) = LcVersao Then Exit Function
   If CLng(LcVerTxt) < LcVersao Then
      Kill LcArq
      Open LcArq For Output As #LcNumero
      Write #LcNumero, LcVersao
      Close #LcNumero
      FileCopy App.Path & "\" & LcNomeExecutavel, LcDir & LcNomeExecutavel '"comercial.exe"
      DoEvents
      Exit Function
   End If
   If CLng(LcVerTxt) > LcVersao Then
      Shell App.Path & "\atualizar.exe " & App.Path & "\" & LcNomeExecutavel & "," & LcDir & LcNomeExecutavel & "," & LcNomeExecutavel, vbNormalFocus
      End
   End If
Else
   Open LcArq For Output As #LcNumero
   Write #LcNumero, LcVersao
   Close #LcNumero
   FileCopy App.Path & "\" & LcNomeExecutavel, LcDir & LcNomeExecutavel '"comercial.exe"
   Exit Function
End If



End Function
Public Function Verificatb()
Dim LcArq   As Integer
Dim LcEnd   As String
Dim LcTab   As String
Dim LcTotal As Integer
Dim w       As Integer

On Error GoTo ErroCriacao

LcEnd = App.Path & "\principal.def"

'=== Verifica se Existe Tabelas para Atualizar
If Dir$(LcEnd) = "" Then Exit Function
'=== Pergunta ao Usuario se Quer aTualizar a base
LcResposta = MsgBox("Existe uma Atualização para ser feita na sua base de dados." & Chr(13) & "Para efetuar esta atualização, Todas as Máquinas deverão estar fora do Sistema." & Chr(13) & "Efetua a Atualização Agora ?", vbInformation + vbYesNo, "Atualização do Sistema.")
If LcResposta = 7 Then Exit Function

'=== mostra a Tela de Aviso
ExibeMsgAtualizacao.Show
DoEvents
'===Abre o Arquivo Principal Para Saber o Nome das Tabelas

LcArq = FreeFile
Open LcEnd For Input As #LcArq
Do Until EOF(LcArq)
    Line Input #LcArq, LcTab
    LcTotal = LcTotal + 1
    DoEvents
Loop
Close #LcArq
w = 1
Open LcEnd For Input As #LcArq
Do Until EOF(LcArq)
    ExibeMsgAtualizacao.Caption = "Atualizando Tabela Nº " & w & " de " & LcTotal
    
    Line Input #LcArq, LcTab
    Triagem (LcTab)
    Call ApagaArq(LcTab)
    DoEvents
    w = w + 1
Loop

SaiVe:
Close #LcArq
Kill LcEnd

Unload ExibeMsgAtualizacao
Exit Function

ErroCriacao:
If err.Number = 3422 Then
   Msg = "Banco de dados Está sendo Usado por outro Usuario." & Chr(133) & "Continua a Processar as outras Tabelas ?"
End If

LcResposta = MsgBox(Msg, vbInformation + vbYesNo, err.Description & err.Number)
If LcResposta = 7 Then
   GoTo SaiVe
Else
   Resume Next
End If


MsgBox err.Description & "  " & err.Number
Resume Next
End Function
Public Function ApagaArq(LcTabela As String)
On Error Resume Next

Dim LcNomeApagar    As String
LcNomeApagar = App.Path & "\" & LcTabela & ".def"
Kill LcNomeApagar
LcNomeApagar = App.Path & "\" & LcTabela & "ind.def"
Kill LcNomeApagar

End Function

Public Function Triagem(LcTabela As String)
On Error GoTo errotria
Dim LcTabelaAchada  As String
Dim LcAchou         As Boolean
Dim a               As Integer

'GLBase = "D:\dadoscli\teste.mdb"
Set Dbbase = OpenDatabase(GLBase, False, False) ' "dBASE III;")
LcAchou = False
'=== Checa se a tabela Exite no Banco de dados
With Dbbase
     For a = .TableDefs.Count - 1 To 0 Step -1
         DoEvents
         If UCase(.TableDefs(a).Name) = UCase(LcTabela) Then
            LcAchou = True
            Exit For
         End If
     Next
End With
If LcAchou Then
   '=== O Banco de Dados Foi Achado, Chama a Atualização
   '=== dos Campos
   Call VerificaCampos(LcTabela)
Else
   '=== A Tabela Nã Foi Achada, Então Chama a Criação da
   '=== Mesma
   Call CriaTabela(LcTabela)
End If
'===  Acerta os indices das tabelas
Call AcertaIndices(LcTabela)
saitria:
Dbbase.Close
Set Dbbase = Nothing
Exit Function

errotria:
If err.Number = 3422 Then
   Msg = "Banco de dados Está sendo Usado por outro Usuario." & Chr(133) & "Continua a Processar as outras Tabelas ?"
Else
   Msg = "Foi encontrado erro no Processamento." & Chr(13) & "Continua o Processamento mesmo Assim?"
End If

LcResposta = MsgBox(Msg, vbInformation + vbYesNo, "Aviso")
If LcResposta = 7 Then
   GoTo errotria
Else
   Resume Next
End If



End Function
Public Function CriaTabela(LcTabela As String)
On Error GoTo errocria

Dim b               As Integer
Dim LcArqTb         As Integer
Dim LcNomeTb        As String
Dim LcNomeCampo     As String
Dim LcEndArqTb      As String
Dim LcLinha         As String
Dim LcTipo          As String
Dim LcTamanho       As String
Dim LCLEtra         As String
Dim LcRequerido     As String
Dim LcZero          As String
Dim LcPos           As Integer
Dim x               As Integer
Dim LcAchouCampo    As Boolean
Dim RsTab           As TableDef

Set RsTab = Dbbase.CreateTableDef(LcTabela)
'Set tdfNew = dbsNorthwind.CreateTableDef("Contacts")
    
'=== Abre o arquivo que contem os dados do campo
LcArqTb = FreeFile
LcEndArqTb = App.Path & "\" & LcTabela & ".def"
Open LcEndArqTb For Input As #LcArqTb

Do Until EOF(LcArqTb)
   DoEvents
   Line Input #LcArqTb, LcLinha
   '=== Separa os dasdos Recebidos
   LcPos = 1
   LcNomeCampo = ""
   LcTipo = ""
   LcTamanho = ""
   LcRequerido = ""
   LcZero = ""
   For x = 1 To Len(LcLinha)
       DoEvents
       LCLEtra = Mid(LcLinha, x, 1)
       If LCLEtra <> "," Then
          Select Case LcPos
              Case Is = 1
                 LcNomeCampo = LcNomeCampo & LCLEtra
              Case Is = 2
                 LcTipo = LcTipo & LCLEtra
              Case Is = 3
                 LcTamanho = LcTamanho & LCLEtra
              Case Is = 4
                 LcRequerido = LcRequerido & LCLEtra
              Case Is = 5
                 LcZero = LcZero & LCLEtra
          End Select
       Else
            LcPos = LcPos + 1
       End If
   Next
   RsTab.Fields.Append RsTab.CreateField(LcNomeCampo, CInt(LcTipo), CInt(LcTamanho))
 '== Acerta a Definicao dos campos
   If CInt(LcTipo) <> 4 Then
      If UCase(LcZero) = "FALSE" Or UCase(LcZero) = "FALSO" Then
         RsTab.Fields(LcNomeCampo).AllowZeroLength = False
      Else
         RsTab.Fields(LcNomeCampo).AllowZeroLength = True
      End If
      '== Acerta o campo Requerido
      If UCase(LcRequerido) = "FALSE" Or UCase(LcRequerido) = "FALSO" Then
         RsTab.Fields(LcNomeCampo).Required = False
      Else
         RsTab.Fields(LcNomeCampob).Required = True
      End If
   End If
Loop
saicria:
Dbbase.TableDefs.Append RsTab
Close #LcArqTb
Exit Function

errocria:
If err.Number = 3422 Then
   Msg = "Banco de dados Está sendo Usado por outro Usuario." & Chr(133) & "Continua a Processar as outras Tabelas ?"
Else
   Msg = "Foi encontrado erro no Processamento." & Chr(13) & "Continua o Processamento mesmo Assim?"
End If

LcResposta = MsgBox(Msg, vbInformation + vbYesNo, "Aviso")
If LcResposta = 7 Then
   GoTo saicria
Else
   Resume Next
End If


End Function
Public Function VerificaCampos(LcTabela As String)
On Error GoTo errocampos

Dim b               As Integer
Dim LcArqTb         As Integer
Dim LcNomeTb        As String
Dim LcNomeCampo     As String
Dim LcEndArqTb      As String
Dim LcLinha         As String
Dim LcTipo          As String
Dim LcTamanho       As String
Dim LCLEtra         As String
Dim LcRequerido     As String
Dim LcZero          As String
Dim LcNomeTab       As String
Dim LcPos           As Integer
Dim x               As Integer
Dim LcAchouCampo    As Boolean
Dim RsTab           As TableDef

Set RsTab = Dbbase.TableDefs(LcTabela)

'=== Abre o arquivo que contem os dados do campo
LcArqTb = FreeFile
LcEndArqTb = App.Path & "\" & LcTabela & ".def"
Open LcEndArqTb For Input As #LcArqTb

Do Until EOF(LcArqTb)
   DoEvents
   Line Input #LcArqTb, LcLinha
   '=== Separa os dasdos Recebidos
   LcPos = 1
   LcPos = 1
   LcNomeCampo = ""
   LcTipo = ""
   LcTamanho = ""
   LcRequerido = ""
   LcZero = ""
   For x = 1 To Len(LcLinha)
       DoEvents
       LCLEtra = Mid(LcLinha, x, 1)
       If LCLEtra <> "," Then
          Select Case LcPos
              Case Is = 1
                 LcNomeCampo = LcNomeCampo & LCLEtra
              Case Is = 2
                 LcTipo = LcTipo & LCLEtra
              Case Is = 3
                 LcTamanho = LcTamanho & LCLEtra
              Case Is = 4
                 LcRequerido = LcRequerido & LCLEtra
              Case Is = 5
                 LcZero = LcZero & LCLEtra
          End Select
       Else
            LcPos = LcPos + 1
       End If
   Next
   LcAchouCampo = False
   With RsTab
     For b = .Fields.Count - 1 To 0 Step -1
         DoEvents
         If UCase(.Fields(b).Name) = UCase(LcNomeCampo) Then
            '== O Campo Existe
            LcAchouCampo = True
            Exit For
         End If
      Next
      
      If Not LcAchouCampo Then
        '=== Vai criar um novo campo
        
        LcNomeTab = .Name
        .Fields.Append .CreateField(LcNomeCampo, CInt(LcTipo), CInt(LcTamanho))
      Else
        '=== Acerta os dados do campo
        '.Fields(LcNomeCampo).Type = CInt(LcTipo)
        '.Fields(LcNomeCampo).Size = CInt(LcTamanho)
      End If
      '== Acerta a opcao comprimento zero
      If .Fields(LcNomeCampo).Type <> 4 Then
         If UCase(LcZero) = "FALSE" Or UCase(LcZero) = "FALSO" Then
            RsTab.Fields(LcNomeCampo).AllowZeroLength = False
         Else
            RsTab.Fields(LcNomeCampo).AllowZeroLength = True
         End If
      '== Acerta o campo Requerido
         If UCase(LcRequerido) = "FALSE" Or UCase(LcRequerido) = "FALSO" Then
            RsTab.Fields(LcNomeCampo).Required = False
         Else
            RsTab.Fields(LcNomeCampob).Required = True
         End If
     End If
   End With
   RsTab.Fields.Refresh
Loop
saicampos:
Close #LcArqTb
Exit Function
errocampos:
If err.Number = 3422 Then
   'Resume Next
   Msg = "Banco de dados Está sendo Usado por outro Usuario." & Chr(133) & "Continua a Processar as outras Tabelas ?"
Else
   Msg = "Foi encontrado erro no Processamento." & Chr(13) & "Continua o Processamento mesmo Assim?"
End If

LcResposta = MsgBox(Msg, vbInformation + vbYesNo, "Erro Nº" & err.Number & "  " & err.Description & "Tb:" & LcNomeTab)
If LcResposta = 7 Then
   GoTo saicampos
Else
   Resume Next
End If


End Function

Public Function AcertaIndices(LcTabela)
On Error GoTo erroacerta

Dim LcArqInd        As Integer
Dim LcEndArqInd     As String
Dim LcLinha         As String
Dim LCLEtra         As String
Dim LcNomeindice    As String
Dim LcCampos        As String
Dim LcRequerido     As String
Dim LcUnico         As String
Dim LcPrimario      As String
Dim a               As Integer
Dim LcApaga         As Integer
Dim LcPos           As Integer
Dim RsTab           As TableDef

Set RsTab = Dbbase.TableDefs(LcTabela)

'=== Apaga os indices Atuais
With RsTab
   For LcApaga = .Indexes.Count - 1 To 0 Step -1
     DoEvents
     .Indexes.Delete .Indexes(LcApaga).Name
   Next
End With

'=== Abre o arquivo que contem os dados do indice
LcArqInd = FreeFile
LcEndArqInd = App.Path & "\" & LcTabela & "ind.def"
Open LcEndArqInd For Input As #LcArqInd

Do Until EOF(LcArqInd)
   DoEvents
   Line Input #LcArqInd, LcLinha
   LcPos = 1
   LcNomeindice = ""
   LcCampos = ""
   LcPrimario = ""
   LcRequerido = ""
   LcUnico = ""
   For a = 1 To Len(LcLinha)
       DoEvents
       LCLEtra = Mid(LcLinha, a, 1)
       If LCLEtra <> "+" Then
          If LCLEtra <> "," Then
             Select Case LcPos
                Case Is = 1
                    LcNomeindice = LcNomeindice & LCLEtra
                Case Is = 2
                    LcCampos = LcCampos & LCLEtra
                Case Is = 3
                    LcPrimario = LcPrimario & LCLEtra
                Case Is = 4
                    LcRequerido = LcRequerido & LCLEtra
                Case Is = 5
                    LcUnico = LcUnico & LCLEtra
              End Select
           Else
             LcPos = LcPos + 1
           End If
       End If
   Next
   '=== Cria os indices das tabelas
   Set LcIndices = RsTab.CreateIndex(LcNomeindice)
   LcIndices.Fields = LcCampos
   If UCase(LcPrimario) = "FALSO" Or UCase(LcPrimario) = "FALSE" Then
      LcIndices.Primary = False
   Else
      LcIndices.Primary = True
   End If
   
   If UCase(LcRequerido) = "FALSO" Or UCase(LcRequerido) = "FALSE" Then
      LcIndices.Required = False
   Else
      LcIndices.Required = True
   End If
   
   If UCase(LcUnico) = "FALSO" Or UCase(LcUnico) = "FALSE" Then
      LcIndices.Unique = False
   Else
     LcIndices.Unique = True
   End If
   '=== Atribui o Indice a tabela
   RsTab.Indexes.Append LcIndices
   RsTab.Indexes.Refresh
   
Loop
SaidaSistema:
Close #LcArqInd
Exit Function

erroacerta:
If err.Number = 3422 Then
   Msg = "Banco de dados Está sendo Usado por outro Usuario." & Chr(133) & "Continua a Processar as outras Tabelas ?"
Else
   Msg = "Foi encontrado erro no Processamento." & Chr(13) & "Continua o Processamento mesmo Assim?"
End If

LcResposta = MsgBox(Msg, vbInformation + vbYesNo, err.Number & err.Description)
If LcResposta = 7 Then
   GoTo SaidaSistema
Else
   Resume Next
End If

End Function
