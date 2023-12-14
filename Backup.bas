Attribute VB_Name = "Backup"

Public Enum eResultado
            Sucesso
            erro
End Enum

Public Enum eDrive
            a
            b
End Enum

Private Busca               As eResultado
Private mvarBackupDrive     As eDrive
Private mvarPastaTemp       As String

'===========================================================================

'=== PARA MONITORAR A EXECUÇÂO DOS PROCESSOS ===============================

Private Type STARTUPINFO
   cb               As Long
   lpReserved       As String
   lpDesktop        As String
   lpTitle          As String
   dwX              As Long
   dwY              As Long
   dwXSize          As Long
   dwYSize          As Long
   dwXCountChars    As Long
   dwYCountChars    As Long
   dwFillAttribute  As Long
   dwFlags          As Long
   wShowWindow      As Integer
   cbReserved2      As Integer
   lpReserved2      As Long
   hStdInput        As Long
   hStdOutput       As Long
   hStdError        As Long
End Type

Private Type PROCESS_INFORMATION
   hProcess         As Long
   hThread          As Long
   dwProcessID      As Long
   dwThreadID       As Long
End Type

Private Declare Function WaitForSingleObject _
Lib "kernel32" _
( _
   ByVal hHandle As Long, _
   ByVal dwMilliseconds As Long _
) As Long

Private Declare Function CreateProcessA _
Lib "kernel32" _
( _
   ByVal lpApplicationName As Long, _
   ByVal lpCommandLine As String, _
   ByVal lpProcessAttributes As Long, _
   ByVal lpThreadAttributes As Long, _
   ByVal bInheritHandles As Long, _
   ByVal dwCreationFlags As Long, _
   ByVal lpEnvironment As Long, _
   ByVal lpCurrentDirectory As Long, _
   lpStartupInfo As STARTUPINFO, _
   lpProcessInformation As PROCESS_INFORMATION _
) As Long

Private Declare Function CloseHandle _
Lib "kernel32" _
( _
   ByVal hObject As Long _
) As Long

Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&
Private Const WAIT_FAILED = &HFFFFFFFF
Private Const WAIT_TIMEOUT = &H102&
Private Const STILL_ACTIVE = &H103&
'===========================================================================

'=== PARA DESCOBRIR O ESPAÇO NO DISQUETE ===================================
Private Declare Function GetDiskFreeSpace Lib "kernel32" _
Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, _
lpSectorsPerCluster As Long, lpBytesPerSector As Long, _
lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters _
As Long) As Long
'===========================================================================

'=== PARA CÓPIA DE ARQUIVO =================================================
Private Type SHFILEOPSTRUCT
  hWnd                  As Long
  wFunc                 As Long
  pFrom                 As String
  pTo                   As String
  fFlags                As Integer
  fAnyOperationsAborted As Boolean
  hNameMappings         As Long
  lpszProgressTitle     As String
End Type

Private Declare Function SHFileOperation Lib _
       "Shell32.dll" Alias _
       "SHFileOperationA" (lpFileOp As _
       SHFILEOPSTRUCT) As Long
       
Private Const FO_COPY = &H2
Private Const FOF_ALLOWUNDO = &H40
'===========================================================================

Private Function Copiar(Origem As String, Destino As String) As eResultado
    'ESTA FUNÇÃO USA A API SHFileOperation PARA COPIAR UM ARQUIVO
    'DE UM LUGAR PARA OUTRO.


    Dim lResultado As Long
    Dim Arquivo As SHFILEOPSTRUCT
    
    'preramos os parâmetros...
    With Arquivo
      .hWnd = 0
      .wFunc = FO_COPY
      'Arquivos a serem copiados separados por NULO
      'e terminado por 2 NULOS
      .pFrom = Origem & vbNullChar & vbNullChar
      'ou, para copiar TODOS os arquivos, use a linha abaixo:
      '.pFrom = "C:\*.*" & vbNullChar & vbNullChar
      'O diretório de destino, ou o nome do arquivo de destino:
      .pTo = Destino & vbNullChar & vbNullChar
      .fFlags = FOF_ALLOWUNDO
    End With
    
    'executamos a cópia...
    lResultado = SHFileOperation(Arquivo)
    If lResultado <> 0 Then
      Copiar = erro
      'exibe a descrição do erro...
      MsgBox Err.LastDllError, vbCritical Or vbOKOnly
    Else
      If Arquivo.fAnyOperationsAborted <> 0 Then
        Copiar = erro
        'avisa que não foi copiado...
        MsgBox "Falha na operação de cópia!", vbCritical Or vbOKOnly
      End If
    End If
    
End Function

Private Function sTamanhoDoDrive(sDrive As String, Optional sLista As String, Optional ByRef RetornoTotalBytes As Long) As String
    'ESTA FUNÇÃO DESCOBRE TAMANHO E ESPAÇO LIVRE DO DRIVE SOLICITADO
    'USANDO A API DO WINDOWS GetDiskFreeSpace
    
    'ESTA FUNÇÃO ESPERA QUE O DISQUETE ESTEJA NO LUGAR E ESTEJA
    'COM O SEU ESPAÇO TOTAL LIVRE PARA USO

    Dim SectorsPerCluster       As Long
    Dim BytesPerSector          As Long
    Dim NumberOfFreeClusters    As Long
    Dim TotalNumberOfClusters   As Long
    Dim BytesFree               As Long
    Dim BytesTotal              As Long
    Dim PercentFree             As Long
    Dim lRetorno                As Long
    Dim FreeBytes               As Long
    Dim TotalBytes              As Long
    Dim ListaBytes              As Long
    
    'chama a API passando as variáveis por referência...
    lRetorno = GetDiskFreeSpace(sDrive, SectorsPerCluster, BytesPerSector, NumberOfFreeClusters, TotalNumberOfClusters)
    'calcula o tamanho do drive...
    TotalBytes = TotalNumberOfClusters * SectorsPerCluster * BytesPerSector
    
    If RetornoTotalBytes = -1 Then
        RetornoTotalBytes = TotalBytes
        Exit Function
    End If
    'calcula o espaço livre...
    FreeBytes = NumberOfFreeClusters * SectorsPerCluster * BytesPerSector
    
        
        
        'o disquete não foi colocado...
        If TotalBytes = 0 Then
            sTamanhoDoDrive = "O Drive " & sDrive & " está vazio. Coloque um disquete formatado e limpo. Para continuar clique em OK"
            Exit Function
        End If
    
        'se o disquete está no lugar, mais está cheio...
        If TotalBytes > 0 And FreeBytes = 0 Then
            sTamanhoDoDrive = "O o Disquete do Drive " & sDrive & " está cheio. Coloque um disquete formatado e limpo. Para continuar clique em OK"
            Exit Function
        End If
    
    
    If sLista <> "" Then
        'se é para copiar a lista...
        
        'descobre o tamanho do arquivo de lista
        ListaBytes = FileLen(sLista)
        
        'o disquete está no lugar mais não há espaço suficiente...
        If TotalBytes < (ListaBytes + FreeBytes) Then
            sTamanhoDoDrive = "O Disquete do Drive " & sDrive & " não tem espaço suficiente. Coloque um disquete formatado e limpo. Para continuar clique em OK"
            Exit Function
        End If
    Else
        'se é para copiar arquivos CAB...
        
        'o disquete está no lugar mais está parcialmente ocupado...
        If TotalBytes > FreeBytes Then
            sTamanhoDoDrive = "O Disquete do Drive " & sDrive & " não tem espaço suficiente. Coloque um disquete formatado e limpo. Para continuar clique em OK"
            Exit Function
        End If
    End If
End Function
Public Sub Descomprimir(Caminho As String)
    'ESTE PROCEDIMENTO DESCOMPACTA ATRAVES DO PROGRAMA EXTRACT.EXE
    'MULTIPLOS ARQUIVOS CAB QUE FORAM CRIADOS COM O PROCEDIMENTO
    'DE COMPACTAÇAO DESTE PROGRAMA.
    'E' NECESSARIO PASSAR O DIRETÓRIO ONDE SE ENCONTRAM OS
    'ARQUIVOS CRIADOS.  DESCOMPACTAÇAO SE DARÁ NO MESMO DIRETÓRIO.
    
    'IMPORTANTE: OS PROGRAMAS MAKECAB.EXE E EXTRACT.EXE
    'DEVERAO SER COPIADOS PARA A PASTA DESTE PROJETO
    'PARA QUE O MESMO POSSA FUNCIONAR
    
    Dim sComando        As String   'linha de comando do programa Extract.exe
    Dim lRetorno        As Long     'Retorno da função Shell
    Dim sArquivo        As String
    Dim ProcInfo        As PROCESS_INFORMATION
    Dim StartProc       As STARTUPINFO
    
    'caminho não informado...
    If Caminho = "" Then
        MsgBox "Não foi informada a pasta de descompressão.", vbOKOnly + vbCritical, "Erro"
        Exit Sub
    End If
    
    'caminho não existe...
    If Dir$(Caminho, vbDirectory) = "" Then
        MsgBox "A pasta informada não existe.", vbOKOnly + vbCritical, "Erro"
        Exit Sub
    Else
        'se o caminho existe, formatamos o caminho para uma subpasta..
        'se a sub pasta não existe, criamos
        '*** (este passo é importante, por que evita sobreescrever
        'arquivos existentes na pasta escolhida) ***
        Caminho = sFormataCaminho(Caminho) & "CAB\"
        If Dir$(Caminho, vbDirectory) = "" Then
            'cria pasta de destino...
            MkDir Caminho
        End If
    End If
    
    'tentamos a recuperação de arquivos do drive escolhido...
    If RecuperarArquivos(Caminho) = erro Then
        Exit Sub
    End If
    
    'formata caminho do primeiro arquivo
    sArquivo = sFormataCaminho(Caminho) & "Backup1.CAB"
    
    'primeiro arquivo não existe...
    If Dir$(sArquivo) = "" Then
        MsgBox "Os arquivos para descompactação não estão na pasta informada", vbOKOnly + vbCritical, "Erro"
        Exit Sub
    End If
    
    'inicializa variável
    Busca = Sucesso
    
    'monta linha de comando...
    sComando = sCommand & sExtract & Caminho & " " & sArquivo
    
    'se houve erro na montagem da linha de comando, aborta...
    If Busca = erro Then Exit Sub
    On Error Resume Next
    
    'executamos a linha de comando e fazemos o monitoramento de sua execução...
    StartProc.cb = Len(StartProc)
    If CreateProcessA(0&, sComando, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, _
                      0&, 0&, StartProc, ProcInfo) Then
    
       '(by Chato de Galocha) * Está como INFINITE mas pode ser trocado para um valor
       'em milisegundos e lRetorno testado com WAIT_FAILED, WAIT_TIMEOUT e STILL_ACTIVE *
       lRetorno = WaitForSingleObject(ProcInfo.hProcess, INFINITE)
       CloseHandle (ProcInfo.hThread)
       CloseHandle (ProcInfo.hProcess)
       
       'avisa do sucesso...
       MsgBox "Os arquivos foram descompactados na pasta " & Caminho & ".", vbOKOnly + vbInformation, "Sucesso"
    Else
       'por algum motivo o processo não foi iniciado... avisa...
       MsgBox "Erro ao executar a descompactação", vbOKOnly + vbCritical, "Erro"
    End If

End Sub

Public Sub Comprimir(ByVal Backup As Boolean, ByVal Caminho As String, Arquivo() As String)
    'ESTA PROCEDURE RECEBE COMO PARAMETROS O DIRETORIO DE TRABALHO
    'E A LISTA DE ARQUIVOS A SEREM COMPCTADOS, MONTA ARQUIVO DE
    'DEFINIÇÃO, MONTA LINHA DE COMANDO PARA PROGRAMA MAKECAB.EXE
    'E EXECUTA A LINHA DE COMANDO.
    
    'IMPORTANTE: OS PROGRAMAS MAKECAB.EXE E EXTRACT.EXE
    'DEVERAO SER COPIADOS PARA A PASTA DESTE PROJETO
    'PARA QUE O MESMO POSSA FUNCIONAR
    

    Dim sCaminhoCAB     As String   'arquivo de definição
    Dim sCaminhoARQ     As String   'linha de comando
    Dim iArquivo        As Integer  'FreeFile
    Dim iNum            As Integer  'Loop pelo Array de Parametros
    Dim lRetorno        As Long     'Retorno da função Shell
    Dim sPathCab        As String   'Destino dos arquivos CAB
    Dim sTemp           As String
    Dim ProcInfo        As PROCESS_INFORMATION
    Dim StartProc       As STARTUPINFO
    
    If Dir$(Caminho, vbDirectory) = "" Then
        MsgBox "O Caminho informado não existe!", vbOKOnly + vbCritical, "Erro"
        Exit Sub
    End If
    
    'captura numero de arquivo livre...
    iArquivo = FreeFile
    
    'captura pasta temp e formata contrabarra
    sTemp = sFormataCaminho(sPastaTemp)
    
    'formata caminho da pasta de destino
    sPathCab = sTemp & "CAB"
    
    'formata caminho do arquivo de definição...
    sCaminhoCAB = sTemp & "CAB.DDF"
    
    'se a pasta de destino não existir, cria...
    If Dir$(sPathCab, vbDirectory) = "" Then
        MkDir sPathCab
    Else
        'prevenimos o erro, caso não existam arquivos a deletar
        On Error Resume Next
        'se já existir, limpa seu conteúdo...
        Kill sPathCab & "\*.*"
    End If
    
    'inicilizamos a variável...
    Busca = Sucesso
    
    'cria arquivo de definição...
    'este arquivo é necessario para o programa MakeCab.exe
    Open sCaminhoCAB For Output As #iArquivo
        Print #iArquivo, ".Option EXPLICIT"
        Print #iArquivo, ".Set Cabinet = off"
        Print #iArquivo, ".Set Compress = off"
        Print #iArquivo, ".Set MaxDiskSize = 1457664"
        Print #iArquivo, ".Set ReservePerCabinetSize = 6144"
        Print #iArquivo, ".Set DiskDirectoryTemplate = " & Chr(34) & sPathCab & Chr(34)
        Print #iArquivo, ".Set CompressionType = MSZIP"
        Print #iArquivo, ".Set CompressionLevel = 7"
        Print #iArquivo, ".Set CompressionMemory = 21"
        Print #iArquivo, ".Set CabinetNameTemplate =" & Chr(34) & "Backup*.CAB" & Chr(34)
        Print #iArquivo, ".Set Cabinet=on"
        Print #iArquivo, ".Set Compress=on"
        
        'percorre a lista de arquivos a serem compactados...
        For iNum = LBound(Arquivo()) To UBound(Arquivo())
            'escreve o nome deste arquivo...
            Print #iArquivo, Chr(34) & Arquivo(iNum) & Chr(34)
        Next
        'fecha arquivo...
    Close #iArquivo
    
    'alterna a pasta atual para a informada...
    ChDir Caminho
    
    'monta linha de comando...
    sCaminhoARQ = sCommand & sMakeCab & sCaminhoCAB
    MsgBox sCaminhoARQ
    'se na moontagem da linha de comando houvem erros, abortamos...
    If Busca = erro Then Exit Sub
    
    On Error Resume Next
    
    'executamos a linha de comando e fazemos o monitoramente de sua execução...
    StartProc.cb = Len(StartProc)
    If CreateProcessA(0&, sCaminhoARQ, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, _
                      0&, 0&, StartProc, ProcInfo) Then
    
       '(by Chato de Galocha) * Está como INFINITE mas pode ser trocado para um valor
       'em milisegundos e lRetorno testado com WAIT_FAILED, WAIT_TIMEOUT e STILL_ACTIVE *
       lRetorno = WaitForSingleObject(ProcInfo.hProcess, INFINITE)
       CloseHandle (ProcInfo.hThread)
       CloseHandle (ProcInfo.hProcess)
       
       'se é para fazer backup...
       If Backup = True Then
         'chamamos a rotina, passando o caminho como parâmetro...
         If CopiarArquivos(sPathCab) = Sucesso Then
             'se o backup for feito até o final, avisa o sucesso...
             MsgBox "Arquivos transferidos com sucesso", vbOKOnly + vbInformation, "Sucesso"
         End If
       Else
         'não era pra fazer backup, então avisa onde estão os arquivos...
         MsgBox "Os arquivos compactados foram criados na pasta " & sPathCab & ".", vbOKOnly + vbInformation, "Sucesso"
       End If
    Else
        'por algum motivo o processo não foi iniciado... avisa...
       MsgBox "Erro ao executar a compactação", vbOKOnly + vbCritical, "Erro"
    End If
End Sub

Private Function sFormataCaminho(ByVal sCaminho As String) As String

    'Verifica se existe "\" no caminho do arquivo
    If Not Right(sCaminho, 1) = Chr(92) Then
        sCaminho = sCaminho & Chr(92)
    End If
    sFormataCaminho = sCaminho
    
End Function

Public Property Let BackupDrive(ByVal vData As eDrive)
    'seleciona drive de backup.
    'o drive default o Drive é A:\ e se outro
    'não for informado, este será usado.
    mvarBackupDrive = vData
End Property


Public Property Get BackupDrive() As eDrive
    'informa Drive de Backup...
    Set BackupDrive = mvarBackupDrive
End Property




Public Property Let PastaTemp(ByVal vData As String)
    'determina uma pasta temp para este processo...
    mvarPastaTemp = vData
End Property


Public Property Get PastaTemp() As String
    'captura a pasta temp escolhida...
    PastaTemp = mvarPastaTemp
End Property




Private Function sPastaTemp() As String
    'CAPTURA A LOCALIZAÇÃO DA PASTA TEMP DO SISTEMA
    'SE O USUÁRIO NÃO INFORMAR UMA
    
    'se o usuário informou a pasta temporária...
    If mvarPastaTemp <> "" Then
        'se ela ainda não existe...
        If Dir$(mvarPastaTemp, vbDirectory) = "" Then
            'criamos...
            MkDir mvarPastaTemp
        End If
    End If
    
    'se a pasta não foi informada...
    If mvarPastaTemp = "" Then
        'tentamos capturar a temporária do sistema...
        mvarPastaTemp = Environ("TMP")
    End If
    
    'se não encontramos anteriormente...
    If mvarPastaTemp = "" Then
        'tentamos novamente...
        mvarPastaTemp = Environ("TEMP")
    End If
    
    'caramba!! não encontramos...
    'na boa, seu computador está ligado? (ehehehehehe....)
    If mvarPastaTemp = "" Then
        'formatamos um caminho...
        mvarPastaTemp = "C:\Temp"
        'se ele ainda não existe...
        If Dir$(mvarPastaTemp, vbDirectory) = "" Then
            'criamos...
            MkDir mvarPastaTemp
        End If
    End If
    
    'informamos...
    sPastaTemp = mvarPastaTemp
    
End Function

Private Function sCommand() As String
    'CAPTURA A LOCALIZAÇÃO DO PROGRAMA COMMAND.COM E MONTA LINHA DE COMANDO
    
    Dim sTemp       As String
    
    'captura a variável de ambiente...
    sTemp = Environ("COMSPEC")
    
    'se por algum motivo (improvável) não está aqui...
    If sTemp = "" Then
        'vamos
        sTemp = Environ("WINDIR")
        
        'capturamos o caminho do Windows
        If sTemp <> "" Then
            'montamos caminho...
            sTemp = sTemp & "Command\command.com"
            'se não está neste caminho (improvável)...
            If Dir$(sTemp) = "" Then
                'montamos caminho...
                sTemp = "C:\command.com"
                
                'caramba!!! não é possível que você não tenha o command.com...
                'Que sistema operacional você usa? (ehehehehe.... )
                If Dir$(sTemp) = "" Then
                    'sinaliza erro na busca...
                    Busca = erro
                    'avisa...
                    MsgBox "Programa Command.com não encontrado", vbOKOnly + vbCritical, "Erro"
                    Exit Function
                End If
            End If
        End If
    End If
      
    'montamos linha de comando
    sCommand = sTemp & " /C "
    
End Function

Private Function sMakeCab() As String
    'CAPTURA A LOCALIZAÇÃO DO PROGRAMA MAKECAB.EXE E MONTA LINHA DE COMANDO

    Dim sTemp       As String
    
    'formata caminho de onde deveria estar...
    sTemp = sFormataCaminho(App.Path) & "Makecab.exe"
    
    'se não está no caminho do aplicativo...
    If Dir$(sTemp) = "" Then
        'sinaliza do erro na busca...
        Busca = erro
        'avisa...
        MsgBox "Programa Makecab.exe não encontrado", vbOKOnly + vbCritical, "Erro"
        Exit Function
    End If
    
    'monta linha de comando...
    sMakeCab = sTemp & " /F "

End Function


Private Function sExtract() As String
    'CAPTURA A LOCALIZAÇÃO DO PROGRAMA EXTRACT.EXE E MONTA LINHA DE COMANDO
    
    Dim sTemp       As String
    
    'formata caminho de onde deveria estar...
    sTemp = sFormataCaminho(App.Path) & "Extract.exe"
    
    'se não está no caminho do aplicativo...
    If Dir$(sTemp) = "" Then
        'sinaliza do erro na busca...
        Busca = erro
        'avisa...
        MsgBox "Programa Extract.exe não encontrado", vbOKOnly + vbCritical, "Erro"
        Exit Function
    End If
    
    'monta linha de comando...
    sExtract = sTemp & " /Y /A /E /L "

End Function

Private Function CopiarArquivos(sPathCab As String) As eResultado
'ESTA FUNÇÃO PREPARA A LISTA DE ARQUIVOS A SEREM COPIADOS PARA DISQUETE
'E CHAMA A FUNÇÃO DE CÓPIA. NESTE PROCESSO, ELA MONITORA O DRIVE PARA
'DESCOBRIR SE O DISQUETE ESTÁ PRONTO E PODERÁ SUPORTAR O TAMANHO DO ARQUIVO.
'ESTA ROTINA ESTÁ PROGRAMADA PARA ACEITAR UM BACKUP DE ATÉ 99 ARQUIVOS CAB
'A SEREM COPIADOS, O QUE JÁ SERIA UM EXAGERO. DIFICILMENTE UM BACKUP ACIMA
'DISTO SERIA FEITO VIA DISQUETE.

    Dim iNum        As Integer
    Dim sTemp       As String
    Dim lResultado  As Long
    Dim sMensagem   As String
    Dim sDrive      As String
    
    'captura escolha de drive para
    'backup. Se o usuário não indicar,
    'o A:\ será usado por padrão
    If mvarBackupDrive = b Then
        sDrive = "B:\"
    Else
        sDrive = "A:\"
    End If
    
    'acerta contrabarra no caminho...
    sPathCab = sFormataCaminho(sPathCab)
    
    'faz um loop pelos possíveis arquivos...
    For iNum = 1 To 99
        'formata nome atual...
        sTemp = sPathCab & "Backup" & iNum & ".CAB"
        'se ele existir, tentamos copiar...
        If Dir$(sTemp) <> "" Then
            'avisa para prepara o disquete...
            MsgBox "Coloque um disquete (número " & iNum & ") formatado e vazio na unidade " & sDrive & " . Aperte OK quando estiver pronto.", vbOKOnly + vbInformation, "Disquete"
LerDrive:
            'verifica o estado do disquete...
            sMensagem = sTamanhoDoDrive(sDrive)
            'se há mensagem é porque o disquete não está pronto...
            If sMensagem <> "" Then
                'se deseja abortar a operação, saimos...
                If MsgBox(sMensagem, vbOKCancel + vbInformation, "Disquete") = vbCancel Then
                    CopiarArquivos = erro
                    Exit Function
                Else
                    'tentamos ler um disquete novamente...
                    GoTo LerDrive
                End If
            Else
                'copiamos o arquivo e capturamos o sucesso...
               If Copiar(sTemp, sDrive) <> Sucesso Then
                    CopiarArquivos = erro
                    Exit Function
               End If
            End If
        Else
            'não existem mais arquivos a copiar, então saimos...
            Exit For
        End If
    Next
    If CopiarLista(iNum, sPathCab, sDrive) = erro Then Exit Function
    'cópias realizadas com sucesso...
    CopiarArquivos = Sucesso
    
End Function


Private Function CopiarLista(Itens As Integer, sPastaLista As String, sDrive As String) As eResultado
Dim iArquivo    As Integer
Dim iNum        As Integer
Dim sLista      As String
Dim sMensagem   As String

sLista = sPastaLista & "ListaCab.txt"
iArquivo = FreeFile
Open sLista For Output As #iArquivo
  Print #iArquivo, "*** ARQUIVO DE DEFINIÇÃO ****"
  Print #iArquivo, "-- não altere este arquivo --"
  Print #iArquivo, "-----------------------------"
    For iNum = 1 To Itens - 1
      Print #iArquivo, "Backup" & iNum & ".CAB"
    Next
  Print #iArquivo, "-----------------------------"
Close #iArquivo

LerDrive:
sMensagem = sTamanhoDoDrive(sDrive, sLista)
If sMensagem <> "" Then
      'se deseja abortar a operação, saimos...
      If MsgBox(sMensagem, vbOKCancel + vbInformation, "Disquete") = vbCancel Then
        CopiarLista = erro
        Exit Function
      Else
        GoTo LerDrive
      End If
End If

If Copiar(sLista, sDrive) = erro Then Exit Function

CopiarLista = Sucesso
End Function

Private Function RecuperarArquivos(Caminho As String) As eResultado
    'ESTA FUNÇÃO RECUPERA OS ARQUIVOS COMPACTADOS QUE DEVERÃO ESTAR NOS DISQUETES
    'O PRIMEIRO ARQUIVO A SER LIDO É O ARQUIVO DE DEFINIÇÃO ONDE CONSTA A LISTA
    'DE ARQUIVOS ASER RECUPERADA
    
    Dim sDrive          As String
    Dim iItem           As Integer
    Dim iArquivo        As Integer
    Dim sArquivo        As String
    Dim sItem           As String
    Dim iNum            As Integer
    Dim sLinha          As String
    Dim lRetorno        As Long
    
    'captura escolha de drive para
    'backup. Se o usuário não indicar,
    'o A:\ será usado por padrão
    If mvarBackupDrive = b Then
        sDrive = "B:\"
    Else
        sDrive = "A:\"
    End If
    'formata nome do arquivo de definição
    sArquivo = sDrive & "ListaCab.txt"
PegaLista:
     lRetorno = -1
     sTamanhoDoDrive sDrive, "", lRetorno
    'se não há disquete no drive
    If lRetorno = 0 Then GoTo Solicita
    'se o arquivo de de definição não foi encontrado...
    If Dir$(sArquivo) = "" Then
Solicita:
        'verificamos se é para continuar...
        If MsgBox("Coloque no drive " & sDrive & "último disque ou o que contém o arquivo de definição ListaCab.txt. Pressione OK para continuar.", vbOKCancel + vbInformation, "Definição") = vbCancel Then
        
            'desistiu... vamos embora...
            RecuperarArquivos = erro
            Exit Function
        Else
            'vamos ler outro disquete...
            GoTo PegaLista
        End If
    End If
    
    'captura número de arquivo livre...
    iArquivo = FreeFile
    
    'abre arquivo...
    Open sArquivo For Input As #iArquivo
        'faz um loop por todo ele, linha alinha...
        Do While Not EOF(iArquivo)
          'lê esta linha...
          Line Input #iArquivo, sLinha
          'contamos as ocorrencias de nomes de arquivos...
          If Mid(sLinha, 1, 6) = "Backup" Then
            'incrementa contador...
            iItem = iItem + 1
          End If
        Loop
    Close #iArquivo
    
    'se algum nome de arquivo foi encontrado...
    If iItem > 0 Then
    
        On Error Resume Next
        '=======================================================================
        '** (CUIDADO = OS ARQUIVOS QUE ESTIVEREM NESTA PASTA SERÃO EXCLUÍDOS) **
        '
        'exclui os arquivos anteriores...
        Kill Caminho & "*.*"
        '
        '=======================================================================
        
        'faz loop pelo número de arquivos esperado...
        For iNum = 1 To iItem
        
            'formata nome do arquivo atual...
            sItem = "Backup" & iNum & ".CAB"
            
CopiarProximo:
            'se o arquivo esperado não foi encontrado...
            If Dir$(sDrive & sItem) = "" Then
                'solicitamos que o coloque no drive...
                If MsgBox("Coloque o disquete número " & iNum & " ou que contenha o arquivo " & sItem & " no drive " & sDrive & " clique em OK.") = vbCancel Then
                    'desistiu...
                  RecuperarArquivos = erro
                  Exit Function
                Else
                    'tentamos ler outro disquete...
                  GoTo CopiarProximo
                End If
            Else
                'tentamos copiar o arquivo CAB do drive para a pasta escolhida...
                If Copiar(sDrive & sItem, Caminho & sItem) = erro Then
                    'se houve um erro, abortamos...
                  RecuperarArquivos = erro
                  Exit Function
                End If
            End If
        Next
    End If

End Function


