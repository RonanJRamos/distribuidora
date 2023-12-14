Attribute VB_Name = "AcessoAdo"
Option Explicit

Public GlUser               As String
Public GlSenhaBase          As String
Public GlPassword           As String
Public LcComentario         As String
Public strConnect           As String
Public GlDirRel             As String
Public LcExibemsg           As Boolean
Public LcRegistrosAfetados  As Long
Public conexaoAdo           As ADODB.Connection
Private LcResp              As Integer
'Public RsAtual As ADODB.Recordset

Function abreconexao()
On Error GoTo erroabrindoconexao
'Dim strConnect As String
Dim TipoConexao As String
Dim Arqini As String
Dim BaseAcess As String
Dim GlDataBase  As String
Dim a As Integer
'esta é a string de conexao devera conter a informacao sobre o provedor e o caminho do banco de dados

Dim strProvider As String
'guarda o nome do provedor
Dim strDataSource As String
'guarda a fonte de dados
Dim strDataBaseName As String
'nome do banco de dados

Dim GlUser As String       ' identificacao do usuario para o banco de dados
Dim GlPassword As String         ' a senha (se tiver) para o banco de dados
Dim GlIp As String    ' o endereco ip da maquina na qual esta o mySql
'==> Determina o nome do arquivo ini
    GlIp = "192.168.0.1" 'LeIni("Base de Dados", "Servidor MySql", Arqini)   ' a localizacao do usuario (localhost)
    GlUser = "root" 'LeIni("base de dados", "usuariosql", Arqini) ' identificacao
    GlPassword = "arquivox" 'LeIni("base de dados", "senhasql", Arqini)      ' senha
    GlDataBase = "lidis" ' LeIni("base de dados", "nomebancosql", Arqini)
 
    ' string de conexao
    strConnect = "driver={MySQL ODBC 3.51 Driver};server=" & GlIp & ";uid=" & GlUser & ";pwd=" & GlPassword & ";database=" & GlDataBase
    'GLBase = buscadirBaseDados
    'GlDirRel = buscadirBaseDados & "rel"

LcComentario = "- Função 'abreconexao - Criando uma Nova Conexão com o DB."
Set conexaoAdo = New ADODB.Connection
'preparando o objeto connection

conexaoAdo.CursorLocation = adUseServer ' adUseClient
'usamos um cursor do lado do cliente pois os dados 'serao acessados na maquina do cliente e nao de um servidor
LcComentario = "- Função 'abreconexao - Abrindo a Conexão com o DB."
'MsgBox strConnect
conexaoAdo.Open strConnect
'conexaoAdo.CommitTrans
Exit Function
erroabrindoconexao:
MsgBox Err.Description & Err.Number
logErro Err.Number, Err.Description, LcComentario & " Provider=" & strConnect
LcResp = ProcessaErro(Err.Number)
If LcResp = 4 Then
   Resume 0
Else
   'End
End If
End Function
Function AbreRecordset(LcSql As String, RsAtual As ADODB.Recordset) As ADODB.Recordset

On Error GoTo ErroAbreRs
LcComentario = "- AbreRecordset - Criando Nova Instancia do RecordSet."
Set RsAtual = New ADODB.Recordset
LcComentario = "- AbreRecordset - Setando os Parametros do Recordset."
RsAtual.CursorType = adOpenDynamic ' adOpenStatic
RsAtual.CursorLocation = adUseServer ' adUseClient
RsAtual.LockType = adLockOptimistic ' adLockPessimistic
'MsgBox LcSql
RsAtual.Source = LcSql
RsAtual.ActiveConnection = conexaoAdo
'Debug.Print LcSql
LcComentario = "- AbreRecordset - Abrindo o Recordset."
RsAtual.Open
Set AbreRecordset = RsAtual
Exit Function

ErroAbreRs:
'If err.Number = 3709 Then
'   'abreconexao
'   Resume 0
'End If
'If LcExibemsg Then ErrosSistema = MsgBox(msg, 64, "erro Abrindo Tabela. ") Else ErrosSistema = 0
MsgBox Err.Description & Err.Number
'Resume 0
logErro Err.Number, Err.Description, LcComentario
LcResp = ProcessaErro(Err.Number, Err.Description)
If LcResp = 4 Then
   Resume 0
Else
   Resume Next
End If
End Function
Function AbreRecordsetLeitura(LcSql As String) As ADODB.Recordset

On Error GoTo ErroAbreRs
Dim RsAtual As ADODB.Recordset
LcComentario = "- AbreRecordset - Criando Nova Instancia do RecordSet."
Set RsAtual = New ADODB.Recordset
LcComentario = "- AbreRecordset - Setando os Parametros do Recordset."
RsAtual.CursorType = adOpenDynamic ' adOpenStatic
RsAtual.CursorLocation = adUseClient
RsAtual.LockType = adLockReadOnly ' adLockOptimistic ' adLockPessimistic

RsAtual.ActiveConnection = conexaoAdo
'LcSql = Replace(LcSql, "*", "%")
'MsgBox conexaoAdo.ConnectionString
RsAtual.Source = LcSql
'Debug.Print LcSql
LcComentario = "- AbreRecordset - Abrindo o Recordset."
RsAtual.Open
Set AbreRecordsetLeitura = RsAtual
Exit Function

ErroAbreRs:
'If err.Number = 3709 Then
'   'abreconexao
'   Resume 0
'End If
'If LcExibemsg Then ErrosSistema = MsgBox(msg, 64, "erro Abrindo Tabela. ") Else ErrosSistema = 0
MsgBox Err.Description & Err.Number
'Resume 0
logErro Err.Number, Err.Description, LcComentario
LcResp = ProcessaErro(Err.Number, Err.Description)
If LcResp = 4 Then
   Resume 0
Else
   Resume Next
End If

End Function
Function VinculaDados(Rs As ADODB.Recordset, lcform As Form)
On Error GoTo errvincula
Dim C       As Control
Dim LcNome  As String
Dim LcMask  As String
Dim LcType  As Integer
Dim LcCodigo As Long
LcComentario = "-Vinculadados - Atualizando o Recordset, com Requery."
LcCodigo = Rs.Fields("codigo")
'Rs.Requery
LcComentario = "-Vinculadados - Voltando para o Registro Atual."
Rs.Find "codigo=" & LcCodigo
LcComentario = "-Vinculadados - Iniciando o Loop no Formulario para obter os Campos."
For Each C In lcform.Controls()
    LcNome = UCase(C.Name)
    '=> Verifica se não é os que não Interessa
    If LcNome <> "NOMERECEITA" And LcNome <> "NOME" And LcNome <> "TITULO" And LcNome <> "BOTOES1" And LcNome <> "BARSTATUS" And LcNome <> "LINE" And LcNome <> "LABEL" And LcNome <> "TAB" And LcNome <> "BOTOES" And LcNome <> "FIGURAS" Then
       LcComentario = "-Vinculadados - Verificando o tipo dos Campos."

       LcType = Rs.Fields(LcNome).Type
       Select Case LcType
            Case Is = 135
                LcComentario = "-Vinculadados - Formatando a data do campo " & LcNome & "."
                If IsDate(Rs.Fields(LcNome)) Then C.Text = Format(Rs.Fields(LcNome), "dd/mm/yy") Else C.Text = "  /  /  "
            Case Is = adDBDate
                LcComentario = "-Vinculadados - Formatando a data do campo " & LcNome & "."
                If IsDate(Rs.Fields(LcNome)) Then C.Text = Format(Rs.Fields(LcNome), "dd/mm/yy") Else C.Text = "  /  /  "
            
            Case adDate
                LcComentario = "-Vinculadados - Formatando a data do campo " & LcNome & "."
                If IsDate(Rs.Fields(LcNome)) Then C.Text = Format(Rs.Fields(LcNome), "dd/mm/yy") Else C.Text = "  /  /  "
            Case Else
                If Len(Rs.Fields(LcNome)) = 0 Or IsNull(Rs.Fields(LcNome)) Then
                   On Error Resume Next
                   LcComentario = "-Vinculadados - Verifiando a máscara do campo " & LcNome & "."
                   LcMask = C.Mask
                   If Err.Number > 0 Then
                      C.Text = Rs.Fields(LcNome) & ""
                   Else
                      LcComentario = "-Vinculadados - Setando o campo com máscara o campo " & LcNome & "."
                      Err.Number = 0
                      LcMask = C.Mask
                      LcMask = Replace(LcMask, "9", " ")
                      C.Text = LcMask
                   End If
                Else
                   C.Text = Rs.Fields(LcNome) & ""
                End If

        End Select
       On Error GoTo errvincula
    End If
Next
Exit Function
errvincula:
'If LcExibemsg Then ErrosSistema = MsgBox(msg, 64, lctitulo) Else ErrosSistema = 0
MsgBox Err.Description & Err.Number
logErro Err.Number, Err.Description, LcComentario
'Resume 0
LcResp = ProcessaErro(Err.Number)
If LcResp = 4 Then
   Resume 0
Else
   End
End If


End Function
Public Function FechaConexao()
On Error Resume Next
LcComentario = "- FechaConexão - Fechando o Recordset."
'RsAtual.Close
LcComentario = "- FechaConexão - Fechando o Banco."
conexaoAdo.Close

LcComentario = "- FechaConexão - Limpando o Recordset."
'Set RsAtual = Nothing
LcComentario = "- FechaConexão - Limpando o Banco."
'set conexaoado = nothing


Exit Function
errfecha:
logErro Err.Number, Err.Description, LcComentario
LcResp = ProcessaErro(Err.Number)
If LcResp = 4 Then
   Resume 0
Else
   End
End If


End Function
Public Function ExcluirRegistro(lcform As Form, Codigo As Long, Rs As ADODB.Recordset) As Boolean
On Error GoTo errorexcluir
Dim LcSql As String
Dim LcRegistros As Integer
Dim LcNome As String
On Error Resume Next
Err.Number = 0
LcNome = Rs!Nome & ""
If Err.Number <> 0 Then
   Err.Number = 0
   LcNome = Rs!Descricao & ""
   If Err.Number <> 0 Then
      Err.Number = 0
      LcNome = Rs!razaosoc & ""
      If Err.Number <> 0 Then
         Err.Number = 0
         LcNome = Rs!Endereco & ""
      End If
         If Err.Number <> 0 Then
            Err.Number = 0
            LcNome = Rs!Nf & ""
            If Err.Number <> 0 Then
               Err.Number = 0
               LcNome = Rs!XTPMONET & ""
               If Err.Number <> 0 Then
                  Err.Number = 0
                  LcNome = Rs!numnf & ""
                  If Err.Number <> 0 Then
                     Err.Number = 0
                     LcNome = Rs!cheque & ""
                     If Err.Number <> 0 Then
                        Err.Number = 0
                        LcNome = Rs!doc & ""
                     End If
                  End If
               End If
            End If
         End If
   End If
End If
On Error GoTo errorexcluir
LcComentario = "-ExcluirRegistro- Setando a Sql para a Exclusão."
LcSql = "delete from " & lcform.Name & " where codigo=" & Codigo

LcComentario = "-ExcluirRegistro- Excluindo o Registro."
conexaoAdo.Execute LcSql, LcRegistrosAfetados
LcComentario = "-ExcluirRegistro- Gravando Log."
Call GravaLogSistema(lcform.Name, "EXCLUSAO", Codigo, LcNome)

LcComentario = "-ExcluirRegistro- Atualizando os Registros."
Rs.Requery
ExcluirRegistro = True
Saida:
Exit Function
errorexcluir:
logErro Err.Number, Err.Description, LcComentario
'MsgBox err.Description & err.Number
ExcluirRegistro = False
LcResp = ProcessaErro(Err.Number)
If LcResp = 4 Then
   Resume 0
Else
   End
End If

End Function
Public Function GravaLogSistema(LcTabela As String, LcAcao As String, LcCodigoReg As Long, LcNome As String)
Dim LcSql As String
On Error Resume Next
LcSql = "Insert into logsistema ("
LcSql = LcSql & "Tabela,nome,codigoregistro,data,hora,maquina,usuario,acao) "
LcSql = LcSql & " values ("
LcSql = LcSql & "'" & LcTabela & "',"
LcSql = LcSql & "'" & LcNome & "',"
LcSql = LcSql & LcCodigoReg & ","

LcSql = LcSql & "'" & Format(Date, "yy-mm-dd") & "',"
LcSql = LcSql & "'" & Format(Time, "hh:mm") & "',"
LcSql = LcSql & "'" & GlNomeMaquina & "',"
LcSql = LcSql & "'" & GlUsuario & "',"
LcSql = LcSql & "'" & LcAcao & "')"
conexaoAdo.Execute LcSql, LcRegistrosAfetados


End Function

Public Function IncluirRegistros(lcform As Form, Rs As ADODB.Recordset) As Boolean
On Error GoTo ErrInclusao
Dim C           As Control
Dim LcNome      As String
Dim LcType      As Integer
Dim LcSql       As String
Dim LcCampos    As String
Dim LcValores   As String
Dim LcCodigo    As String
Dim LcNomeL     As String
Dim LcIncluiCampo As Boolean
LcComentario = "-IncluirRegistros- Gerando a SQL para a inclusão."
LcSql = ""
LcSql = "Insert into " & lcform.Name
LcComentario = "-IncluirRegistros- Efetualndo o Loop no Form para obter os campos e Valores."
For Each C In lcform.Controls()
    LcNome = UCase(C.Name)
    LcIncluiCampo = False
    If LcNome <> "NOMERECEITA" And LcNome <> "NOME" And LcNome <> "TITULO" And LcNome <> "CODIGO" And LcNome <> "BOTOES1" And LcNome <> "BARSTATUS" And LcNome <> "LINE" And LcNome <> "LABEL" And LcNome <> "TAB" And LcNome <> "BOTOES" And LcNome <> "FIGURAS" Then
        LcComentario = "-IncluirRegistros- Separando os tipos do campo."
        LcType = Rs.Fields(LcNome).Type
        Select Case LcType
          Case 135
                 LcComentario = "-IncluirRegistros- Setando tipo de Data."
                 If IsDate(C.Text) Then
                    If Len(LcValores) > 0 Then LcValores = LcValores & ","
                    LcValores = LcValores & "#" & Format(C.Text, "mm/dd/yy") & "#"
                    LcIncluiCampo = True
                 End If

           Case adDate
                 LcComentario = "-IncluirRegistros- Setando tipo de Data."
                 If IsDate(C.Text) Then
                    If Len(LcValores) > 0 Then LcValores = LcValores & ","
                    LcValores = LcValores & "#" & Format(C.Text, "mm/dd/yy") & "#"
                    LcIncluiCampo = True
                 End If
            Case Is = adDBDate
                 LcComentario = "-IncluirRegistros- Setando tipo de Data."
                 If IsDate(C.Text) Then
                    If Len(LcValores) > 0 Then LcValores = LcValores & ","
                    LcValores = LcValores & "'" & Format(C.Text, "yy-mm-dd") & "'"
                    LcIncluiCampo = True
                 End If
            Case Is = dbBoolean
                 LcComentario = "-IncluirRegistros- Setando tipo Boleano."
                  If Len(LcValores) > 0 Then LcValores = LcValores & ","
                  If CInt(C.Value) = 0 Then LcValores = LcValores & "False" Else LcValores = LcValores & "True"
                  LcIncluiCampo = True
            Case adInteger
                 LcComentario = "-IncluirRegistros- Setando tipo Inteiro."
                 If IsNumeric(C.Text) Then
                    If Len(LcValores) > 0 Then LcValores = LcValores & ","
                    LcValores = LcValores & C.Text
                    LcIncluiCampo = True
                 End If
            
            Case Is = adNumeric
                 LcComentario = "-IncluirRegistros- Setando tipo Numerico."
                 If IsNumeric(C.Text) Then
                    If Len(LcValores) > 0 Then LcValores = LcValores & ","
                    C.Text = Replace(C.Text, ",", ".")
                    LcValores = LcValores & C.Text
                    LcIncluiCampo = True
                 End If
            Case Is = adDouble
                 LcComentario = "-IncluirRegistros- Setando tipo Numerico."
                 If IsNumeric(C.Text) Then
                    If Len(LcValores) > 0 Then LcValores = LcValores & ","
                    C.Text = Replace(C.Text, ",", ".")
                    LcValores = LcValores & C.Text
                    LcIncluiCampo = True
                 End If
            Case Is = adDecimal
                 LcComentario = "-IncluirRegistros- Setando tipo Numerico."
                 If IsNumeric(C.Text) Then
                    If Len(LcValores) > 0 Then LcValores = LcValores & ","
                    C.Text = Replace(C.Text, ",", ".")
                    LcValores = LcValores & C.Text
                    LcIncluiCampo = True
                 End If
            Case Is = adInteger
                 LcComentario = "-IncluirRegistros- Setando tipo Numerico."
                 If IsNumeric(C.Text) Then
                    If Len(LcValores) > 0 Then LcValores = LcValores & ","
                    C.Text = Replace(C.Text, ",", ".")
                    LcValores = LcValores & C.Text
                    LcIncluiCampo = True
                 End If
            Case Is = adCurrency
                 LcComentario = "-IncluirRegistros- Setando tipo Numerico."
                 If IsNumeric(C.Text) Then
                    If Len(LcValores) > 0 Then LcValores = LcValores & ","
                    C.Text = Replace(C.Text, ",", ".")
                    LcValores = LcValores & C.Text
                    LcIncluiCampo = True
                 End If
            
            Case Is = adLongVarChar
                   LcComentario = "-IncluirRegistros- Setando tipo String."
                   If Len(LcValores) > 0 Then LcValores = LcValores & ","
                    C.Text = Replace(C.Text, ",", " ")
                    LcValores = LcValores & "'" & UCase(C.Text) & "'"
                    LcIncluiCampo = True
            Case Is = adChar
                    LcComentario = "-IncluirRegistros- Setando tipo String."
                    If Len(LcValores) > 0 Then LcValores = LcValores & ","
                    C.Text = Replace(C.Text, ",", " ")
                    LcValores = LcValores & "'" & UCase(C.Text) & "'"
                    LcIncluiCampo = True
                    
            Case Is = adVarChar
                   LcComentario = "-IncluirRegistros- Setando tipo String."
                   If Len(LcValores) > 0 Then LcValores = LcValores & ","
                    C.Text = Replace(C.Text, ",", " ")
                    LcValores = LcValores & "'" & UCase(C.Text) & "'"
                    LcIncluiCampo = True
           
        End Select
        If LcIncluiCampo Then
            LcComentario = "-IncluirRegistros- Incluindo o Campo na Sql."
            
            If Len(LcCampos) > 0 Then
               LcCampos = LcCampos & ","
            End If
            LcCampos = LcCampos & LcNome
        End If
    End If
Next
LcSql = LcSql & " (" & LcCampos & ") values (" & LcValores & ")"
'Debug.Print LcSql
LcComentario = "-IncluirRegistros- Efetuando a Inclusão na Tabela."
'MsgBox LcSql
conexaoAdo.Execute LcSql, LcRegistrosAfetados
LcComentario = "-IncluirRegistros- Atualizando os Registros."

LcComentario = "-AlteraRegistro- Gravando o Log."
LcCodigo = lcform.Codigo.Text & ""
Err.Number = 0
LcNomeL = lcform.Nome.Text & ""
If Err.Number > 0 Then
    Err.Number = 0
    LcNomeL = lcform.razaosoc.Text & ""
    If Err.Number > 0 Then
       Err.Number = 0
       LcNomeL = lcform.razaosoc.Text & ""
    End If
    LcNomeL = lcform.Descricao.Text & ""
    
End If
Call GravaLogSistema(lcform.Name, "INCLUSAO", CLng(LcCodigo), LcNomeL)

IncluirRegistros = True
Saida:
Exit Function
ErrInclusao:
logErro Err.Number, Err.Description, LcComentario

'MsgBox err.Description & err.Number
IncluirRegistros = False
LcResp = ProcessaErro(Err.Number)
If LcResp = 4 Then
   Resume 0
Else
   End
End If

End Function

Function AlteraRegistro(lcform As Form, LcCodigo As Long, Rs As ADODB.Recordset) As Boolean
On Error GoTo ErrAlteracao
Dim C           As Control
Dim LcNome      As String
Dim LcType      As Integer
Dim LcSql       As String
Dim LcCampos    As String
Dim LcValores   As String
Dim LcNomeL     As String
Dim LcPrimeiro  As Boolean
Dim LcIncluiCampo As Boolean

On Error Resume Next
Err.Number = 0
LcNome = Rs!Nome & ""
If Err.Number <> 0 Then
   Err.Number = 0
   LcNome = Rs!Descricao & ""
   If Err.Number <> 0 Then
      Err.Number = 0
      LcNome = Rs!razaosoc & ""
      If Err.Number <> 0 Then
         Err.Number = 0
         LcNome = Rs!Endereco & ""
      End If
         If Err.Number <> 0 Then
            Err.Number = 0
            LcNome = Rs!Nf & ""
            If Err.Number <> 0 Then
               Err.Number = 0
               LcNome = Rs!XTPMONET & ""
               If Err.Number <> 0 Then
                  Err.Number = 0
                  LcNome = Rs!numnf & ""
                  If Err.Number <> 0 Then
                     Err.Number = 0
                     LcNome = Rs!cheque & ""
                     If Err.Number <> 0 Then
                        Err.Number = 0
                        LcNome = Rs!doc & ""
                     End If
                  End If
               End If
            End If
         End If
   End If
End If
LcNomeL = LcNome
On Error GoTo ErrAlteracao
LcPrimeiro = True
LcComentario = "-AlteraRegistro- Criando Sql."
LcSql = ""
LcSql = "Update " & lcform.Name & " SET "
LcComentario = "-AlteraRegistro- Efetuando o Loop No Form para buscar os campos e Valores."
For Each C In lcform.Controls()
    LcNome = UCase(C.Name)
    LcIncluiCampo = False
    If LcNome <> "NOMERECEITA" And LcNome <> "TITULO" And LcNome <> "CODIGO" And LcNome <> "BOTOES1" And LcNome <> "BARSTATUS" And LcNome <> "LINE" And LcNome <> "LABEL" And LcNome <> "TAB" And LcNome <> "BOTOES" And LcNome <> "FIGURAS" Then
        LcComentario = "-AlteraRegistro- Setando o Tipo do Campo."
        LcType = Rs.Fields(LcNome).Type
        Select Case LcType
           Case 135
                 LcComentario = "-AlteraRegistro- Setando o Tipo Data."
                 If IsDate(C.Text) Then
                    Rs(LcNome).Value = Format(C.Text, "dd/mm/YY")
                 Else
                    Rs(LcNome).Value = Null
                 End If

            Case adDate
                 LcComentario = "-AlteraRegistro- Setando o Tipo Data."
                 If IsDate(C.Text) Then
                    Rs(LcNome).Value = Format(C.Text, "dd/mm/YY")
                 Else
                    Rs(LcNome).Value = Null
                 End If
            
            Case Is = adDBDate
                 LcComentario = "-AlteraRegistro- Setando o Tipo Data."
                 If IsDate(C.Text) Then
                    Rs(LcNome).Value = Format(C.Text, "dd/mm/dd")
                 Else
                    Rs(LcNome).Value = Null
                 End If
            Case Is = dbBoolean
                  LcComentario = "-AlteraRegistro- Setando o Tipo Boleano."
                  Rs(LcNome).Value = C.vaslue
            Case Is = adDouble
                 LcComentario = "-AlteraRegistro- Setando o Tipo Inteiro."
                 If IsNumeric(C.Text) Then
                    Rs(LcNome).Value = C.Text
                 End If
            
            Case Is = adDecimal
                 LcComentario = "-AlteraRegistro- Setando o Tipo Inteiro."
                 If IsNumeric(C.Text) Then
                    Rs(LcNome).Value = C.Text
                 End If
            
            Case Is = adInteger
                 LcComentario = "-AlteraRegistro- Setando o Tipo Inteiro."
                 If IsNumeric(C.Text) Then
                    Rs(LcNome).Value = C.Text
                 
                 End If
            
            Case Is = adCurrency
                 LcComentario = "-AlteraRegistro- Setando o Tipo Inteiro."
                 If IsNumeric(C.Text) Then
                    Rs(LcNome).Value = C.Text
                 End If
            
            Case adInteger
                 LcComentario = "-AlteraRegistro- Setando o Tipo Inteiro."
                 If IsNumeric(C.Text) Then
                    Rs(LcNome).Value = C.Text
                 End If
            
            Case Is = adNumeric
                 LcComentario = "-AlteraRegistro- Setando o Tipo Numérico."
                 If IsNumeric(C.Text) Then
                    Rs(LcNome).Value = C.Text
                 End If
            Case Is = adLongVarChar
                   LcComentario = "-AlteraRegistro- Setando o Tipo String."
                   Rs(LcNome).Value = UCase(C.Text)
            Case Is = adChar
                   LcComentario = "-AlteraRegistro- Setando o Tipo String."
                   Rs(LcNome).Value = UCase(C.Text)
                    
            Case Is = adVarChar
                   Rs(LcNome).Value = UCase(C.Text)
        End Select
       
    End If
Next
LcComentario = "-AlteraRegistro- Efetuando a alteração."
Rs.Update
LcComentario = "-AlteraRegistro- Gravando o Log."
Call GravaLogSistema(lcform.Name, "ALTERAÇÂO", CLng(LcCodigo), LcNomeL)

LcComentario = "-AlteraRegistro- Atualizando o recordset."


'Rs.Requery

AlteraRegistro = True
Saida:
Exit Function
ErrAlteracao:
logErro Err.Number, Err.Description, LcComentario
'MsgBox err.Description & err.Number
AlteraRegistro = False
'Resume 0
LcResp = ProcessaErro(Err.Number)
If LcResp = 4 Then
   Resume 0
Else
   End
End If


End Function

Function GeraPainel(lcform As Form, LcAcao As Integer, Rs As ADODB.Recordset)
On Error GoTo errpainel
Dim LcBook As Variant
Dim LcTotal As Long
Dim LcPos As Integer
Dim LcOrdem As String
Dim LcSt As String
LcComentario = "-GeraPainel- Setando o tamanho do Painel."
lcform.BarStatus.Font.Bold = True
lcform.BarStatus.Font.Size = 11
lcform.BarStatus.Panels(1).Width = 2000
lcform.BarStatus.Panels(2).Width = 4500
lcform.BarStatus.Panels(3).Width = 3400
LcComentario = "-GeraPainel- Verificando a Guantidade de Registros."
If Not Rs.EOF Then
    LcBook = Rs.Bookmark
    Rs.MoveLast
    LcTotal = Rs.RecordCount
    Rs.Bookmark = LcBook
Else
    LcTotal = 0
End If
'StatusBar.Panels(1).Text = "Usuário:" & GlUsuario
'If Len(GlNomeMaquina) = 0 Then
'  StatusBar.Panels(2).Text = "Local"
'Else
'  StatusBar.Panels(2).Text = "Máquina:" & GlNomeMaquina
'End If
LcComentario = "-GeraPainel- Verificando a Ordem Atual."

LcOrdem = Rs.Sort

If Len(LcOrdem) > 0 Then
   LcOrdem = UCase(Mid(LcOrdem, 1, 1)) & LCase(Right(LcOrdem, Len(LcOrdem) - 1))
End If
Select Case LcAcao
    Case 1
        LcSt = "Inclusão"
    Case 2
        LcSt = "Alteração"
    Case 3
        LcSt = "Consultar"
End Select
LcComentario = "-GeraPainel- Escrevendo no Painel."
lcform.BarStatus.Font.Size = 8
lcform.BarStatus.Panels(1).Text = "Nº Registros: " & LcTotal
lcform.BarStatus.Panels(2).Text = "Ordem Atual: " & LcOrdem
lcform.BarStatus.Panels(3).Text = "Status Atual: " & LcSt
Exit Function
errpainel:
logErro Err.Number, Err.Description, LcComentario
Resume Next
End Function
Function logErro(LcNumeroerro As String, LcDesc As String, LcComentario As String)
Dim LcRepete, LcIcone As Integer, msg, lctitulo, LcNomeArquivo As String
Dim LcExibemsg As Integer
Dim LcDiretorio As String
Dim LcGrifa     As String
Dim a           As Long
Dim Arqini      As String
Dim TipoConexao As String
Dim LcNumero    As Integer
Arqini = BuscaDirWin
Arqini = Arqini & "\" & App.EXEName & ".ini"
On Error Resume Next
    
LcGrifa = String(80, "-")
'For a = Len(GLBase) To 1 Step -1
'    If Mid(GLBase, a, 1) = "\" Then Exit For
'Next
TipoConexao = LeIni("Base de Dados", "tipo de banco", Arqini)

If TipoConexao = "MySql" Then
   LcDiretorio = buscadirBaseDados 'Mid(GLBase, 1, a)
Else
   LcDiretorio = LeIni("base de dados", "BaseAcess", Arqini)
   For a = Len(LcDiretorio) To 1 Step -1
       If Mid(LcDiretorio, a, 1) = "\" Then Exit For
   Next
   LcDiretorio = Mid(LcDiretorio, 1, a - 1)
End If

LcIcone = 64
LcNumero = FreeFile

LcNomeArquivo = LcDiretorio & "\ErrosSistema.txt"

Open LcNomeArquivo For Append As #LcNumero      ' Open file for output.
 Print #LcNumero, "Data       :" & Date
 Print #LcNumero, "Hora       :" & Time
 Print #LcNumero, "Maquina    :" & GlNomeMaquina
 Print #LcNumero, "Usuário    :" & GlUsuario
 Print #LcNumero, "Descrição  :" & LcDesc
 Print #LcNumero, "Nº do Erro :" & LcNumeroerro
 Print #LcNumero, "Comentario :" & LcComentario
 Print #LcNumero, LcGrifa
Close #LcNumero
'MsgBox LcDesc


End Function

Function buscadirBaseDados() As String
Dim LcArq       As Integer
Dim LcNumero    As Integer
Dim integridade As String
Dim LcPAth      As String
Dim Mt          As Variant
Dim Pos         As Integer
Dim GLBase1     As String
LcArq = FreeFile

LcPAth = BuscaDirWin()
Open LcPAth & "\my.ini" For Input As #LcArq

Err.Number = 0
Do Until EOF(LcArq)
   Input #LcArq, integridade
   Pos = InStr(UCase(integridade), UCase("datadir"))
   If Pos > 0 Then
      Mt = Split(integridade, "=")
      Exit Do
   End If
Loop
If UBound(Mt) > 0 Then
   GLBase1 = Mt(1)
Else
   GLBase1 = ""
End If
buscadirBaseDados = GLBase1
'glbase=replace(GLBase,"/"
End Function
Function LimpaControles(lcform As Form)
On Error Resume Next
Dim LcNome As String
Dim LcMask As String
Dim C As Control
For Each C In lcform.Controls()
    LcNome = UCase(C.Name)
    '=> Verifica se não é os que não Interessa
    If LcNome <> "LABEL" And LcNome <> "TAB" And LcNome <> "BOTOES" Then
       Err.Number = 0
       LcMask = C.Mask
       If Err.Number > 0 Then
          Err.Number = 0
          LcMask = C.Value
          If Err.Number > 0 Then
             C.Text = ""
          Else
             C.Value = 0
          End If
       Else
         LcMask = Replace(LcMask, "9", " ")
         C.Text = LcMask
       End If
     End If
Next

End Function
Public Function ProcessaErro(Optional Nerro As Long = 0, Optional Descricao As String = "") As Integer
On Error Resume Next
Dim LcMsg As String
Dim LcResp As Integer
Dim LcBotao As Integer

LcMsg = "Ocorreu o Segunte Erro:"

LcBotao = vbRetryCancel
Select Case Nerro
    Case Is = 3707, -2146824581
        LcMsg = LcMsg & Chr(13) & "Não é possível alterar a propriedade ActiveConnection de um objeto Recordset que possua um objeto Command como sua origem."
    Case Is = 3732, -2146824556
        LcMsg = LcMsg & Chr(13) & "O servidor não pode concluir a operação."
    Case Is = 3748, -2146824540
        LcMsg = LcMsg & Chr(13) & "Conexão negada. A nova conexão solicitada possui características diferentes da que está em uso."
    Case Is = 3220, -2146825068
        LcMsg = LcMsg & Chr(13) & "O provedor fornecido é diferente do que está em uso."
    Case Is = 3724, -2146824564
        LcMsg = LcMsg & Chr(13) & "O valor dos dados não pode ser convertido por razões diferentes de incompatibilidade de assinaturas ou estouro de dados. Por exemplo, a conversão pode ter dados truncados."
    Case Is = 3725, -2146824563
        LcMsg = LcMsg & Chr(13) & "O valor dos dados não pode ser definido ou recuperado porque o tipo de dados do campo era desconhecido ou o provedor não possuía recursos para efetuar a operação."
    Case Is = 3747, -2146824541
        LcMsg = LcMsg & Chr(13) & "A operação requer um ParentCatalog válido."
    Case Is = 3726, -2146824562
        LcMsg = LcMsg & Chr(13) & "O registro não contém este campo."
    Case Is = 3421, -2146824867
        LcMsg = LcMsg & Chr(13) & "O aplicativo está usando um valor incorreto para a operação atual."
    Case Is = 3721, -2146824567
        LcMsg = LcMsg & Chr(13) & "O valor dos dados é muito grande para ser representado pelo tipo de dados do campo."
    Case Is = 3738, -2146824550
        LcMsg = LcMsg & Chr(13) & "O URL (Uniform Resources Locator, localizador de recursos uniforme) do objeto a ser excluído está fora do escopo do registro atual."
    Case Is = 3750, -2146824538
        LcMsg = LcMsg & Chr(13) & "O provedor não oferece suporte a restrições de compartilhamento."
    Case Is = 3751, -2146824537
        LcMsg = LcMsg & Chr(13) & "O provedor não oferece suporte ao tipo solicitado de restrição de compartilhamento."
    Case Is = 3251 - 2146825037
        LcMsg = LcMsg & Chr(13) & "O objeto ou provedor não é capaz de efetuar a operação solicitada."
    Case Is = 3749, -2146824539
        LcMsg = LcMsg & Chr(13) & "Falha na atualização dos campos. Para obter informações adicionais, examine a propriedade Status dos objetos de campo individuais."
    Case Is = 3219, -2146825069
        LcMsg = LcMsg & Chr(13) & "Operação não permitida neste contexto."
    Case Is = 3719, -2146824569
        LcMsg = LcMsg & Chr(13) & "O valor dos dados está em conflito com as restrições de integridade do campo."
    Case Is = 3246, -2146825042
        LcMsg = LcMsg & Chr(13) & "Um objeto de conexão não pode ser fechado explicitamente durante uma transação."
    Case Is = 3001, -2146825287
        LcMsg = LcMsg & Chr(13) & "Os argumentos são incorretos, estão fora do intervalo aceitável ou estão em conflito."
    Case Is = 3709, -2146824579
        LcMsg = LcMsg & Chr(13) & "Operação não permitida em um objeto com referência a uma conexão fechada ou inválida."
    Case Is = 3708, -2146824580
        LcMsg = LcMsg & Chr(13) & "Objeto Parameter definido incorretamente. As informações são inconsistentes ou incompletas."
    Case Is = 3714, -2146824574
        LcMsg = LcMsg & Chr(13) & "A transação de coordenação é inválida ou ainda não foi iniciada."
    Case Is = 3729, -2146824559
        LcMsg = LcMsg & Chr(13) & "O URL contém caracteres inválidos. Certifique-se de que o URL está digitado corretamente."
    Case Is = 3265, -2146825023
        LcMsg = LcMsg & Chr(13) & "O item não pode ser encontrado na coleção correspondente ao nome ou ao ordinal solicitado."
    Case Is = 3021, -2146825267
        LcMsg = LcMsg & Chr(13) & "As propriedades BOF ou EOF são True, ou o registro atual foi excluído. A operação solicitada pelo aplicativo requer um registro atual."
    Case Is = 3715, -2146824573
        LcMsg = LcMsg & Chr(13) & "A operação não pode ser efetuada enquanto não houver execução."
    Case Is = 3710, -2146824578
        LcMsg = LcMsg & Chr(13) & "A operação não pode ser efetuada durante o processamento de um evento."
    Case Is = 3704, -2146824584
        LcMsg = LcMsg & Chr(13) & "Operação não permitida quando o objeto está fechado."
    Case Is = 3367, -2146824921
        LcMsg = LcMsg & Chr(13) & "O objeto já está na coleção. Não é possível acrescentar."
    Case Is = 3420, -2146824868
        LcMsg = LcMsg & Chr(13) & "O objeto não é mais válido."
    Case Is = 3705, -2146824583
        LcMsg = LcMsg & Chr(13) & "Operação não permitida quando o objeto está aberto."
    Case Is = 3002, -2146825286
       LcMsg = LcMsg & Chr(13) & "Não foi possível abrir o Banco de Dados."
    Case Is = 3712, -2146824576
        LcMsg = LcMsg & Chr(13) & "Operação cancelada pelo usuário."
    Case Is = 3734, -2146824554
        LcMsg = LcMsg & Chr(13) & "A operação não pode ser efetuada. O provedor não pode obter espaço de armazenamento suficiente."
    Case Is = 3720, -2146824568
        LcMsg = LcMsg & Chr(13) & "Permissão insuficiente impede gravação no campo."
    Case Is = 3742, -2146824546
        LcMsg = LcMsg & Chr(13) & "O valor da propriedade está em conflito com uma propriedade relacionada."
    Case Is = 3739, -2146824549
        LcMsg = LcMsg & Chr(13) & "A propriedade não se aplica ao campo especificado."
    Case Is = 3740, -2146824548
        LcMsg = LcMsg & Chr(13) & "Atributo da propriedade inválido."
    Case Is = 3741, -2146824547
        LcMsg = LcMsg & Chr(13) & "Valor da propriedade inválido. Certifique-se de que o valor está digitado corretamente."
    Case Is = 3743, -2146824545
        LcMsg = LcMsg & Chr(13) & "A propriedade é somente leitura ou não pode ser configurada."
    Case Is = 3744, -2146824544
        LcMsg = LcMsg & Chr(13) & "O valor opcional da propriedade não foi definido."
    Case Is = 3745, -2146824543
        LcMsg = LcMsg & Chr(13) & "O valor somente leitura da propriedade não foi definido."
    Case Is = 3746, -2146824542
        LcMsg = LcMsg & Chr(13) & "O provedor não oferece suporte à propriedade."
    Case Is = 3000, -2146825288
        LcMsg = LcMsg & Chr(13) & "Falha no provedor ao executar a operação solicitada."
    Case Is = 3706, -2146824582
        LcMsg = LcMsg & Chr(13) & "Provedor não encontrado. É possível que ele não esteja instalado corretamente."
    Case Is = 3003, -2146825285
        LcMsg = LcMsg & Chr(13) & "Não foi possível ler o arquivo."
    Case Is = 3731, -2146824557
        LcMsg = LcMsg & Chr(13) & "A operação de cópia não pode ser efetuada. O objeto nomeado pelo URL de destino já existe. Especifique adCopyOverwrite para substituir o objeto."
    Case Is = 3730, -2146824558
        LcMsg = LcMsg & Chr(13) & "O banco de dados está bloqueado por um ou mais processos. Aguarde até que o processo tenha sido concluído e tente a operação novamente."
    Case Is = 3735, -2146824553
        LcMsg = LcMsg & Chr(13) & "O URL de origem ou de destino está fora do escopo do registro atual."
    Case Is = 3722, -2146824566
        LcMsg = LcMsg & Chr(13) & "O valor dos dados está em conflito com o tipo de dados ou com as restrições do campo."
    Case Is = 3723, -2146824565
        LcMsg = LcMsg & Chr(13) & "Falha na conversão. O valor dos dados era assinado enquanto o tipo de dados do campo usado pelo provedor não o era."
    Case Is = 3713, -2146824575
        LcMsg = LcMsg & Chr(13) & "A operação não pode ser efetuada durante uma conexão assíncrona."
    Case Is = 3711, -2146824577
        LcMsg = LcMsg & Chr(13) & "A operação não pode ser efetuada durante uma execução assíncrona."
    Case Is = 3728, -2146824560
        LcMsg = LcMsg & Chr(13) & "Permissões insuficientes para acessar a árvore ou subárvore."
    Case Is = 3736, -2146824552
        LcMsg = LcMsg & Chr(13) & "A operação não pôde ser concluída e o status não está disponível. O campo pode estar indisponível ou não houve tentativa de operação."
    Case Is = 3716, -2146824572
        LcMsg = LcMsg & Chr(13) & "As configurações de segurança deste computador proíbem o acesso a uma fonte de dados em outro domínio."
    Case Is = 3727, -2146824561
        LcMsg = LcMsg & Chr(13) & "O URL de origem ou o pai do URL de destino não existe."
    Case Is = 3737, -2146824551
        LcMsg = LcMsg & Chr(13) & "O registro nomeado por este URL não existe."
    Case Is = 3733, -2146824555
        LcMsg = LcMsg & Chr(13) & "O provedor não pode localizar o dispositivo de armazenamento indicado pelo URL. Certifique-se de que o URL está digitado corretamente."
    Case Is = 3004, -2146825284
        LcMsg = LcMsg & Chr(13) & "Falha ao gravar no arquivo."
    Case Is = 3717, -2146824571
        LcMsg = LcMsg & Chr(13) & "Somente para uso interno. Não use."
    Case Is = 3718, -2146824570
        LcMsg = LcMsg & Chr(13) & "Somente para uso interno. Não use."
    Case Else
        LcMsg = LcMsg & Chr(13) & Descricao
End Select
LcMsg = LcMsg & Chr(13) & "O que deseja fazer?"
LcResp = MsgBox(LcMsg, vbCritical + LcBotao, "erro nº:" & Nerro)
'MsgBox err.Description

ProcessaErro = LcResp

End Function
