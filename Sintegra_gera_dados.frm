VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form Sintegra_gera_dados 
   BackColor       =   &H00EFE3C5&
   Caption         =   "Gerar os dados do Sintegra"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   8265
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdGravaResultado 
      Caption         =   "Grava Resultado"
      Height          =   252
      Left            =   3840
      TabIndex        =   5
      Top             =   240
      Width           =   1692
   End
   Begin VB.CommandButton CommandCmdProcessar1 
      Caption         =   "&Processar"
      Height          =   252
      Left            =   2160
      TabIndex        =   4
      Top             =   240
      Width           =   1452
   End
   Begin VB.ListBox Processados 
      Appearance      =   0  'Flat
      Height          =   4710
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   8055
   End
   Begin MSMask.MaskEdBox Datai 
      Height          =   252
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   852
      _ExtentX        =   1508
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   8
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Dataf 
      Height          =   252
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   852
      _ExtentX        =   1508
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   8
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin VB.Label msg 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   5400
      Width           =   7935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   660
   End
End
Attribute VB_Name = "Sintegra_gera_dados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Tipo50
    icms As String
    valor As Double
End Type

Private Sub CmdGravaResultado_Click()
Dim FnunNota As Integer
FnunNota = FreeFile
LcNota = "c:\NotasProcessadasSintegra.txt"
Open LcNota For Output Access Write As #FnunNota  'Abre Porta Nf
     Print #FnunNota, " Periodo de Processamento:" & Datai.Text & " a " & Dataf.Text
     For a = 0 To Processados.ListCount - 1
          Print #FnunNota, Processados.List(a)
     Next
Close #FnunNota
MsgBox "Arquivo Gerado.", 64, "Aviso"

End Sub

Private Sub CommandCmdProcessar1_Click()
ProcessaSintegra
End Sub

Function ProcessaSintegra()
Dim RsNotas                 As ADODB.Recordset
Dim RsdadosNota             As ADODB.Recordset
Dim Resposta                As String
Dim rsCliente               As Recordset
Dim StrSql                  As String
Dim Total_Notas             As Long
Dim TotalItens              As Long
Dim Notas_Processadas       As Long
Dim Valor_Desconto_Banco    As Double
Dim db                      As Database
Dim Achou                   As Boolean
Dim MT()                    As Tipo50
On Error GoTo erroProcessamento

Set db = OpenDatabase(GLBase)
conexaoAdo.BeginTrans
'==> Abre a tabela de notas fiscais
LcCap = Me.Caption
Me.Caption = "Abrindo a tabela de notas fiscais..."
DoEvents
StrSql = "Select * from alid050 where DTEMIS Between '" & Format(Datai.Text, "yyyy-mm-dd") & "' And '" & Format(Dataf.Text, "yyyy-mm-dd") & "' order by numnf;"
Set RsNotas = AbreRecordset(StrSql)
Me.Caption = "Abrindo a tabela de dados das notas fiscais..."
DoEvents
StrSql = "Select * from alid052 order by numnf,item "
Set RsdadosNota = AbreRecordset(StrSql)

Me.Caption = "Calculando o total de notas a processar..."
DoEvents

If Not RsNotas.EOF Then
   RsNotas.MoveLast
   Total_Notas = RsNotas.RecordCount
   RsNotas.MoveFirst
End If

Do Until RsNotas.EOF
   Notas_Processadas = Notas_Processadas + 1
   Me.Caption = "Abrindo a Tabela de dados dos itens da nota ..."
   DoEvents
   '==> Gravando os dados na tb sintegra
   Me.Caption = "Gravando os dados na tabela sintegra ..."
   DoEvents
    Set rsCliente = db.OpenRecordset("Select * from alid001 where codigo='" & RsNotas!Cliente & "'")
    'MsgBox Txt(8).Text
    If Not rsCliente.EOF Then
       If Not IsNull(rsCliente!CGC) Then
          Cnpj = Replace(rsCliente!CGC, ".", "")
          Cnpj = Replace(Cnpj, ",", "")
          Cnpj = Replace(Cnpj, "-", "")
          Cnpj = Replace(Cnpj, "/", "")
          Cnpj = Replace(Cnpj, "\", "")
          Cnpj = Trim(Cnpj)
          If Len(Cnpj) = 0 Then
            Cnpj = Replace(rsCliente!cpf, ".", "")
            Cnpj = Replace(Cnpj, ",", "")
            Cnpj = Replace(Cnpj, "-", "")
            Cnpj = Replace(Cnpj, "/", "")
            Cnpj = Replace(Cnpj, "\", "")
            Cnpj = Trim(Cnpj)
          End If
       Else
         Cnpj = Replace(rsCliente!Cnpj, ".", "")
         Cnpj = Replace(Cnpj, ",", "")
         Cnpj = Replace(Cnpj, "-", "")
         Cnpj = Replace(Cnpj, "/", "")
         Cnpj = Replace(Cnpj, "\", "")
         Cnpj = Trim(Cnpj)
       End If
       Inscricao = rsCliente!INSCEST
       Inscricao = Replace(Inscricao, ".", "")
       Inscricao = Replace(Inscricao, ",", "")
       Inscricao = Replace(Inscricao, "-", "")
       Inscricao = Replace(Inscricao, "/", "")
       Inscricao = Replace(Inscricao, "\", "")
       Inscricao = Trim(Inscricao)
       Estado = rsCliente!Estado
    End If
    Set rsCliente = Nothing
    
    '==> Insere o cabecalho do Sintegra
    StrSql = "Insert into sintegra (data,nf,cfop,valor,Cliente_Forn,Origem) Values ('" & _
           Format(RsNotas!DTEMIS, "yyyy-mm-dd") & "','" & _
           Right("000000" & RsNotas!numnf, 6) & "','" & _
           Replace(Replace(RsNotas!Cfop, ",", ""), ".", "") & "'," & _
           Replace(Replace(Replace(Replace(Replace(CDbl(RsNotas!ValorNota), ".", ""), ",", "."), "R", ""), "$", ""), " ", "") & "," & _
           RsNotas!Cliente & ",'" & _
           "S" & "')"
    ExecutaSql StrSql
    a = 0
    b = 0
    C = 0
   
   
   '==> Verificando desconto
    Me.Caption = "Calculando o desconto da nota " & RsNotas!numnf
    DoEvents
    If IsNull(RsNotas!Desconto) Then
       Valor_Desconto_Banco = 0
    Else
       Valor_Desconto_Banco = RsNotas!Desconto
    End If
    '==> Calculando o numero de itens da nota
    Valor_Desconto = 0
    If Valor_Desconto_Banco > 0 Then
        RsdadosNota.Find "numnf='" & RsNotas!numnf & "'"
        TotalItens = 0
        Do While RsdadosNota!numnf = RsNotas!numnf
           TotalItens = TotalItens + 1
           RsdadosNota.Find "numnf='" & RsNotas!numnf & "'", 1, adSearchForward
           If RsdadosNota.EOF Then RsdadosNota.MoveFirst
        Loop
        RsdadosNota.MoveFirst
        Valor_Desconto = 0
        If Valor_Desconto_Banco > 0 Then
           Valor_Desconto = CDbl(Valor_Desconto_Banco) / TotalItens
           Valor_Desconto = CDbl(AcertaNumero(CStr(Valor_Desconto), 2))
           
        End If
    Else
      Valor_Desconto = 0
      If RsdadosNota.EOF Then RsdadosNota.MoveFirst
    End If
   RsdadosNota.Find "numnf='" & RsNotas!numnf & "'"
   Do While RsdadosNota!numnf = RsNotas!numnf
        Me.Caption = "Gerando o registro 54 para a NF " & RsNotas!numnf & " Item " & RsdadosNota!Descricao
        DoEvents
        Achou = False
        If b = 0 Then
           ReDim Preserve MT(b)
           MT(b).icms = RsdadosNota!icms
           MT(b).valor = RsdadosNota!QTDE * RsdadosNota!VALUNIT
           b = b + 1
        Else
           For C = 0 To UBound(MT)
              If MT(C).icms = RsdadosNota!icms Then
                 Achou = True
                 Exit For
              End If
           Next
           If Achou Then
                MT(C).valor = CDbl(RsdadosNota!QTDE * RsdadosNota!VALUNIT) + MT(C).valor
           Else
                ReDim Preserve MT(b)
                MT(b).icms = RsdadosNota!icms
                MT(b).valor = i
                b = b + 1
           End If
        End If
        StrSql = "insert into sintegra_54 (cnpj,modelo,serie,nf,cfop,cst,item," & _
                 "codproduto,quantidade,valor_total_bruto,valor_desconto," & _
                 "Base_calculo,base_calculo_subst,ipi,Aliquota_icms,data) Values('" & _
                 Cnpj & "','" & _
                 "01" & "','" & _
                 "1  " & "','" & _
                 Right("000000" & RsNotas!numnf, 6) & "','" & _
                 Replace(Replace(RsNotas!Cfop, ",", ""), ".", "") & "','" & _
                 IIf(IsNull(RsdadosNota!cst), "000", IIf(Len(RsdadosNota!cst) = 0, "000", RsdadosNota!cst)) & "','" & _
                 RsdadosNota!item & "','" & _
                 RsdadosNota!codProd & "'," & _
                 Replace(CDbl(RsdadosNota!QTDE), ",", ".") & "," & _
                 Replace(Replace(Replace(Replace(Replace(CDbl(RsdadosNota!QTDE * RsdadosNota!VALUNIT), ".", ""), ",", "."), "R", ""), "$", ""), " ", "") & "," & _
                 Replace(Replace(Replace(Replace(Replace(CStr(Valor_Desconto), ".", ""), ",", "."), "R", ""), "$", ""), " ", "") & "," & _
                 IIf(CDbl(RsdadosNota!icms) > 0, Replace(Replace(Replace(Replace(Replace(CDbl(CDbl(RsdadosNota!QTDE * RsdadosNota!VALUNIT)) - Valor_Desconto, ".", ""), ",", "."), "R", ""), "$", ""), " ", ""), 0) & "," & _
                 "0" & "," & _
                 "0" & "," & _
                 Replace(RsdadosNota!icms, ",", ".") & ",'" & _
                 Format(RsNotas!DTEMIS) & "')"
                 'MsgBox StrSql
           aferados = ExecutaSql(StrSql)
           RsdadosNota.MoveNext
      Loop
      
      If RsdadosNota.EOF Then
         RsdadosNota.MoveFirst
      Else
         RsdadosNota.MovePrevious
      End If
      Me.Caption = "Gerando o registro 50 para a NF " & RsNotas!numnf
      DoEvents
      '==> Grava o reguistro 50
      For a = 0 To UBound(MT)
            StrSql = "Insert into sintegra_50 (Cnpj,inscricao,data,uf,modelo,serie," & _
                     "nf,cfop,emitente,valortotal,base_calculo_icms,Valor_icms," & _
                     "isenta,outra,aliquota,situacao) values ('" & _
                     Cnpj & "','" & _
                     Inscricao & "','" & _
                     Format(RsNotas!DTEMIS, "yyyy-mm-dd") & "','" & _
                     Estado & "','" & _
                     "01" & "','" & _
                     "1   " & "','" & _
                     Right("000000" & RsNotas!numnf, 6) & "','" & _
                     Replace(Replace(RsNotas!Cfop, ",", ""), ".", "") & "','" & _
                     "P" & "'," & _
                     Replace(MT(a).valor, ",", ".") & "," & _
                     IIf(CDbl(MT(a).icms) > 0, Replace(MT(a).valor, ",", "."), 0) & "," & _
                     Replace(AcertaNumero(CStr((CDbl(MT(a).icms) / 100) * MT(a).valor), 2), ",", ".") & "," & _
                     "0" & "," & _
                     "0" & "," & _
                     Replace(MT(a).icms, ",", ".") & ",'" & _
                     IIf(RsNotas!Status <> "CANCELADA", "N", "S") & "')"
                     'MsgBox StrSql
             ExecutaSql (StrSql)
                     
      Next
      Processados.AddItem RsNotas!numnf & " Processado OK."
      msg.Caption = "Processadas " & Notas_Processadas & " de " & Total_Notas
      DoEvents
      RsNotas.MoveNext
Loop
conexaoAdo.CommitTrans
Set RsNotas = Nothing
Me.Caption = LcCap
MsgBox "Processamento efetuado com sucesso.", 64, "Aviso"

Exit Function
erroProcessamento:
'MsgBox "Ocorreu o seguinte erro processando os dados para o sintegra:" & err.Description & " Nº:" & err.Number & Chr(13) & Chr(13) & "Nenhum dado foi salvo. Corriga o erro e processe novamente.", 64, "Aviso"
'Resume 0
conexaoAdo.RollbackTrans
MsgBox "Ocorreu o seguinte erro processando os dados para o sintegra:" & err.Description & " Nº:" & err.Number & Chr(13) & Chr(13) & "Nenhum dado foi salvo. Corriga o erro e processe novamente.", 64, "Aviso"

End Function

