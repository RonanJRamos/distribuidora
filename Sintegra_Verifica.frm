VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form Sintegra_Verifica 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Verificar Dados do Sintegra"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   9000
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdGerarArquivo 
      Caption         =   "Gerar Arquivo Texto"
      Height          =   375
      Left            =   4800
      TabIndex        =   5
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton CmdProcessar 
      Caption         =   "&Processar"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.ListBox Lista 
      Height          =   4545
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   8415
   End
   Begin MSMask.MaskEdBox Datai 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox DataF 
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   360
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   540
   End
End
Attribute VB_Name = "Sintegra_Verifica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function VerificaCliente(CodigoCliente As String) As String
Dim db As Database
Dim Rs As Recordset
Dim StrSql As String
Dim Retorno As String
Dim Resposta As String
Dim Cpf_Cnpj As String
Dim Estado As String
Dim Inscricao As String

StrSql = "Select * from Alid001 where codigo='" & CodigoCliente & "'"
Set db = OpenDatabase(GLBase)
db.Execute "UPDATE ALID001 SET ALID001.INSCEST = 'ISENTO' WHERE (((ALID001.INSCEST)=''));"
'MsgBox Db.RecordsAffected
Set Rs = db.OpenRecordset(StrSql)

If Rs.EOF Then
   Retorno = "Cliente não encontrado."
Else
   '==> Verifica o Cnpj / Cpf
   If IsNull(Rs!CGC) Then
      Cpf_Cnpj = Rs!cpf & ""
   Else
        Cpf_Cnpj = Rs!CGC & ""
        Cpf_Cnpj = Replace(Cpf_Cnpj, ".", "")
        Cpf_Cnpj = Replace(Cpf_Cnpj, ",", "")
        Cpf_Cnpj = Replace(Cpf_Cnpj, "-", "")
        Cpf_Cnpj = Replace(Cpf_Cnpj, "/", "")
        Cpf_Cnpj = Replace(Cpf_Cnpj, "\", "")
        Cpf_Cnpj = Replace(Cpf_Cnpj, " ", "")
        
        If Len(Cpf_Cnpj) = 0 Then
            Cpf_Cnpj = Rs!cpf & ""
            Cpf_Cnpj = Replace(Cpf_Cnpj, ".", "")
            Cpf_Cnpj = Replace(Cpf_Cnpj, ",", "")
            Cpf_Cnpj = Replace(Cpf_Cnpj, "-", "")
            Cpf_Cnpj = Replace(Cpf_Cnpj, "/", "")
            Cpf_Cnpj = Replace(Cpf_Cnpj, "\", "")
            Cpf_Cnpj = Replace(Cpf_Cnpj, " ", "")
        End If
   End If
   Resposta = Verifica_cpf_Cnpj(Cpf_Cnpj)
   If Len(Resposta) > 0 Then
      Retorno = Retorno & Resposta
   End If
   Estado = ""
   '==> verifica o Estado
   If IsNull(Rs!Estado) Then
      Retorno = Retorno & " Estado não cadastro."
   Else
     If Len(Rs!Estado) = 0 Then
        Retorno = Retorno & " Estado não cadastro."
     Else
        Estado = Rs!Estado
     End If
   End If
   Inscricao = Rs!INSCEST & ""
   
   If Len(Inscricao) > 0 Then
        '==> Verifica a Inscricao
        Inscricao = Replace(Inscricao, ".", "")
        Inscricao = Replace(Inscricao, ",", "")
        Inscricao = Replace(Inscricao, "-", "")
        Inscricao = Replace(Inscricao, "/", "")
        Inscricao = Replace(Inscricao, "\", "")
        Inscricao = Replace(Inscricao, " ", "")
        If Consiste(Inscricao, Estado) <> 0 Then
           Retorno = Retorno & "A Inscrição Estadual do cliente é invalida."
        End If
   Else
      Retorno = Retorno & " Inscricao Estadual não Cadastrada."
   End If
End If

Set Rs = Nothing
If Len(Retorno) > 0 Then
   VerificaCliente = Retorno
Else
   VerificaCliente = Estado
End If
End Function
Function Verifica_cpf_Cnpj(Cnpj_CPf As String) As String
Dim Srtcnpj As String
Dim Resposta As Boolean

Srtcnpj = Replace(Cnpj_CPf, ".", "")
Srtcnpj = Replace(Srtcnpj, ",", "")
Srtcnpj = Replace(Srtcnpj, "-", "")
Srtcnpj = Replace(Srtcnpj, "/", "")
Srtcnpj = Replace(Srtcnpj, "\", "")
Srtcnpj = Replace(Srtcnpj, " ", "")

If Len(Srtcnpj) > 11 Then
   Resposta = Calc_CNPJ(Srtcnpj)
Else
   Resposta = Calc_CPF(Srtcnpj)
End If

If Not Resposta Then
   Verifica_cpf_Cnpj = "Cnpj / Cpf inválido."
End If
End Function


Function VerificaCFOP(Estado As String, Cfop As String) As String

End Function

Private Sub CmdGerarArquivo_Click()
Dim FnunNota As Integer
FnunNota = FreeFile
LcNota = "c:\ArquivosSintegraErro.txt"
Open LcNota For Output Access Write As #FnunNota  'Abre Porta Nf
     For a = 0 To Lista.ListCount - 1
          Print #FnunNota, Lista.List(a)
     Next
Close #FnunNota
MsgBox "Arquivo Gerado.", 64, "Aviso"
End Sub

Private Sub CmdProcessar_Click()
'On Error Resume Next
Dim StrSql As String
Dim RsNotas As ADODB.Recordset
Dim Resposta As String

StrSql = "Update Alid050 set cfop='5102' where cfop='512';"
ExecutaSql StrSql
StrSql = "Update Alid050 set cfop='5905' where cfop='599';"
ExecutaSql StrSql


StrSql = "Select * from alid050 WHERE DTEMIS Between '" & Format(Datai.Text, "yyyy-mm-dd") & "' And '" & Format(Dataf.Text, "yyyy-mm-dd") & "';"
Set RsNotas = AbreRecordset(StrSql)

Do Until RsNotas.EOF
   DoEvents
   Resposta = VerificaCliente(RsNotas!Cliente)
   If Len(Resposta) > 2 Then
      GravaLista "Nota:" & RsNotas!numnf & " Cliente:" & RsNotas!Cliente & " - " & Resposta
   Else
     If UCase(Resposta) = "MG" Then
        If Left(RsNotas!Cfop, 1) <> "5" Then
           RsNotas!Cfop = "5" & Right(RsNotas!Cfop, Len(RsNotas!Cfop) - 1)
           GravaLista "Nota:" & RsNotas!numnf & " Cliente:" & RsNotas!Cliente & " - CFOP Invalido para o cliente dentro do estado."
        End If
     Else
        If Left(RsNotas!Cfop, 1) <> "6" Then
           RsNotas!Cfop = "6" & Right(RsNotas!Cfop, Len(RsNotas!Cfop) - 1)
           GravaLista "Nota:" & RsNotas!numnf & " Cliente:" & RsNotas!Cliente & " - CFOP Invalido para o cliente fora do estado."
        End If
     
     End If
   
   End If
   
   RsNotas.MoveNext
Loop
MsgBox "Processamento Terminado.", 64, "Aviso"

End Sub
Sub GravaLista(Dado As String)
Dim Achou As Boolean
Dim LcPos() As String
Dim LcCli() As String
'==> Verifica se o dado ja foi colocado na lista
Achou = False
LcPos = Split(Dado, " ")
For a = 0 To Lista.ListCount - 1
   LcCli = Split(Lista.List(a), " ")
   If LcCli(1) = LcPos(1) Then
      Achou = True
      Exit Sub
   End If
Next

If Not Achou Then Lista.AddItem Dado
End Sub

