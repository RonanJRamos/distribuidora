Attribute VB_Name = "DisqueteReceitaFederal"
Private Type ValorIcmsDisk
    icms  As String
    valor As Currency
    subst As Currency
End Type
Private Type subst
    valor As Currency
    NF    As String
End Type

Function GeraDisquete(drive As String)
On Error GoTo errgera

Dim Rs As ADODB.Recordset
Dim RsCliente As Recordset
Dim RsEmpresa As Recordset
Dim RsDados As ADODB.Recordset
Dim a As Long
Dim LcAbriu As Boolean

Dim db As Database
Dim FnunBoleto As Integer

Dim LcSql As String
Dim Mt() As ValorIcmsDisk
Dim MtS() As subst

Dim ca As Integer
Dim LcTotalRegistros As Long
Dim LcTotalSubst     As Long
Dim CaS              As Long
FnunBoleto = FreeFile

LcTotalRegistros = 0
Set db = OpenDatabase(GLBase)
Set RsEmpresa = db.OpenRecordset("Select * from Empresa", dbOpenDynaset, dbSeeChanges, dbOptimistic)
ca = 0
CaS = 0
LcSql = "Select * from alid050 where DTEMIS Between '" & Format(DisqueteReceita.Datai.Text, "yyyy-mm-dd") & "' And '" & Format(DisqueteReceita.Dataf.Text, "yyyy-mm-dd") & "'"
'abreconexao
Set Rs = AbreRecordset(LcSql)
LcBoleto = drive & "registrofiscal.txt"
Screen.MousePointer = 11
LcCap = DisqueteReceita.Caption
DisqueteReceita.Caption = "Gerando Registros "

Open LcBoleto For Output Access Write As #FnunBoleto 'Abre Porta Nf
LcAbriu = True

'===> Gera Registro Mestre
LcRegistro = "10"
            
LCCnpj = RsEmpresa!CGC & ""
LCCnpj = Replace(LCCnpj, ".", "")
LCCnpj = Replace(LCCnpj, "-", "")
LCCnpj = Replace(LCCnpj, "/", "")
LCCnpj = Left(LCCnpj & String(14, "0"), 14)
LcRegistro = LcRegistro & LCCnpj
            
LcInsc = RsEmpresa!Inscricao & ""
LcInsc = Replace(LcInsc, ".", "")
LcInsc = Replace(LcInsc, "-", "")
LcInsc = Replace(LcInsc, "/", "")
LcInsc = Left(LcInsc & String(14, " "), 14)
If Len(Trim(LcInsc)) = 0 Then LcInsc = Left("ISENTO" & String(14, " "), 14)
LcRegistro = LcRegistro & LcInsc
LcRegistro = LcRegistro & Left(RsEmpresa!Razao & String(35, " "), 35)
LcRegistro = LcRegistro & Left(RsEmpresa!cidade & String(30, " "), 30)
LcRegistro = LcRegistro & RsEmpresa!Estado

LcFax = Right(RsEmpresa!Fax, 10)
LcFax = Replace(LcFax, "(", "")
LcFax = Replace(LcFax, ")", "")
LcFax = Replace(LcFax, "-", "")
LcFax = Replace(LcFax, ".", "")
LcRegistro = LcRegistro & LcFax

LcRegistro = LcRegistro & Format(DisqueteReceita.Datai.Text, "YYYYMMDD")
LcRegistro = LcRegistro & Format(DisqueteReceita.Dataf.Text, "YYYYMMDD")
LcRegistro = LcRegistro & "   "
Print #FnunBoleto, LcRegistro


Do Until Rs.EOF
   'Set RsCliente = AbreRecordset("select * from alid001 where codigo='" & Rs!CLIENTE & "'", RsCliente)
   Set RsCliente = db.OpenRecordset("select * from alid001 where codigo='" & Rs!cliente & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)
   Set RsDados = AbreRecordset("Select * from alid052 where numnf='" & Rs!numnf & "'")
   ca = 0
   ReDim Mt(ca)
   Do Until RsDados.EOF
      If IsNull(RsDados!icms) Then
          LcIcms = 0
      Else
          If Len(RsDados!icms) = 0 Then
            LcIcms = 0
          Else
            LcIcms = RsDados!icms
          End If
      End If
       
      If ca = 0 Then
               ReDim Preserve Mt(ca)
               Mt(ca).icms = LcIcms
               If CDbl(LcIcms) <> 0 Then
                  Mt(ca).valor = RsDados!qtde * RsDados!VALUNIT
               Else
                 Mt(ca).subst = RsDados!qtde * RsDados!VALUNIT
               End If
               ca = ca + 1
      Else
               LcAchou = False
               For a = 0 To ca - 1
                  If Mt(a).icms = LcIcms Then
                     LcAchou = True
                     Exit For
                  End If
               Next
               If LcAchou Then
                 If CDbl(LcIcms) <> 0 Then
                   Mt(a).valor = Mt(a).valor + (RsDados!qtde * RsDados!VALUNIT)
                 Else
                   Mt(a).subst = Mt(a).subst + (RsDados!qtde * RsDados!VALUNIT)
                 End If
               Else
                  ReDim Preserve Mt(ca)
                  Mt(ca).icms = LcIcms & ""
                  If CDbl(LcIcms) <> 0 Then
                     Mt(ca).valor = RsDados!qtde * RsDados!VALUNIT
                  Else
                     Mt(ca).subst = RsDados!qtde * RsDados!VALUNIT
                  End If
                  ca = ca + 1
               End If
      End If
      
       RsDados.MoveNext
       'ca = 0
   Loop
   '===> Escreve no disquete
   If Not RsCliente.EOF Then
        For a = 0 To UBound(Mt)
            LcRegistro = "50"
            LCCnpj = RsCliente!CGC & ""
            LCCnpj = Replace(LCCnpj, ".", "")
            LCCnpj = Replace(LCCnpj, "-", "")
            LCCnpj = Replace(LCCnpj, "/", "")
            LCCnpj = Left(LCCnpj & String(14, "0"), 14)
            LcRegistro = LcRegistro & LCCnpj
            
            LcInsc = RsCliente!INSCEST & ""
            LcInsc = Replace(LcInsc, ".", "")
            LcInsc = Replace(LcInsc, "-", "")
            LcInsc = Replace(LcInsc, "/", "")
            LcInsc = Left(LcInsc & String(14, " "), 14)
            If Len(Trim(LcInsc)) = 0 Then LcInsc = Left("ISENTO" & String(14, " "), 14)
            LcRegistro = LcRegistro & LcInsc
                   
            
            LcRegistro = LcRegistro & Format(Rs!DTEMIS, "YYYYMMDD")
            
            LcEst = RsCliente!Estado & ""
            If Len(LcEst) = 0 Then LcEst = "MG"
            LcRegistro = LcRegistro & LcEst3
            
            LcRegistro = LcRegistro & "01"
            
            LcRegistro = LcRegistro & "001" '==> Serie
            
            LcRegistro = LcRegistro & "  " '==> Sub Serie
            
            LcRegistro = LcRegistro & Rs!numnf
            
            LcCfpo = Rs!CFOP & ""
            LcCfpo = Replace(LcCfpo, ".", "")
            LcCfpo = Left(LcCfpo & String(4, " "), 4)
            LcRegistro = LcRegistro & LcCfpo
            
            LcValorTotal = AcertaNumero(CStr(Rs!ValorNota), 2)
            LcValorTotal = Replace(LcValorTotal, ".", "")
            LcValorTotal = Replace(LcValorTotal, ",", "")
            LcValorTotal = Right("0000000000000" & LcValorTotal, 13)
            LcRegistro = LcRegistro & LcValorTotal
            
            LcBaseCalculo = AcertaNumero(CStr(Mt(a).valor), 2)
            LcBaseCalculo = Replace(LcBaseCalculo, ".", "")
            LcBaseCalculo = Replace(LcBaseCalculo, ",", "")
            LcBaseCalculo = Right("0000000000000" & LcBaseCalculo, 13)
            LcRegistro = LcRegistro & LcBaseCalculo
            If Len(Mt(a).icms) = 0 Then Mt(a).icms = 0
            LcValorIcms = Mt(a).valor * (CCur(Mt(a).icms) / 100)
            
            LcValorIcms = AcertaNumero(CStr(LcValorIcms), 2)
            LcValorIcms = Replace(LcValorIcms, ".", "")
            LcValorIcms = Replace(LcValorIcms, ",", "")
            LcValorIcms = Right("0000000000000" & LcValorIcms, 13)
            LcRegistro = LcRegistro & LcValorIcms
            
            LcRegistro = LcRegistro & "0000000000000"
            
            LcValorOutras = AcertaNumero(CStr(Mt(a).subst), 2)
            
            LcValorOutras = AcertaNumero(CStr(LcValorOutras), 2)
            LcValorOutras = Replace(LcValorOutras, ".", "")
            LcValorOutras = Replace(LcValorOutras, ",", "")
            LcValorOutras = Right("0000000000000" & LcValorOutras, 13)
            LcRegistro = LcRegistro & LcValorOutras
            
            LcAliquota = AcertaNumero(Mt(a).icms, 2)
                   
            LcAliquota = Replace(LcAliquota, ".", "")
            LcAliquota = Replace(LcAliquota, ",", "")
            LcAliquota = Right("0000" & LcAliquota, 4)
            LcRegistro = LcRegistro & LcAliquota
            
            If Rs!Natureza = "CANCELADA" Then
               LcSituacao = "S"
            Else
               LcSituacao = "N"
            End If
            LcRegistro = LcRegistro & LcSituacao
            LcTotalRegistros = LcTotalRegistros + 1
            Print #FnunBoleto, LcRegistro
            DisqueteReceita.Caption = "Gerando Registros. Nº " & LcTotalRegistros
            DoEvents
        Next
    End If
   Rs.MoveNext
Loop
'===> Fianliza o Registro
LcRegistro = "90" & LCCnpj
LcRegistro = LcRegistro & LcInsc
LcTotal50 = AcertaNumero(CStr(LcTotalRegistros), 2)
LcTotal50 = Replace(LcTotal50, ".", "")
LcTotal50 = Replace(LcTotal50, ",", "")
LcTotal50 = Right("00000000" & LcTotal50, 8)
LcRegistro = LcRegistro & LcTotal50

LcTotal51 = Right("00000000" & LcTotal51, 8)
LcRegistro = LcRegistro & LcTotal51

LcTotal53 = Right("00000000" & LcTotal53, 8)
LcRegistro = LcRegistro & LcTotal53

LcTotal60 = Right("00000000" & LcTotal60, 8)
LcRegistro = LcRegistro & LcTotal60

LcTotal61 = Right("00000000" & LcTotal61, 8)
LcRegistro = LcRegistro & LcTotal61

LcTotal70 = Right("00000000" & LcTotal70, 8)
LcRegistro = LcRegistro & LcTotal70

LcTotal71 = Right("00000000" & LcTotal71, 8)
LcRegistro = LcRegistro & LcTotal71

LcTotalGeral = AcertaNumero(CStr(CDbl(LcTotalRegistros) + 2), 2)
LcTotalGeral = Replace(LcTotalGeral, ".", "")
LcTotalGeral = Replace(LcTotalGeral, ",", "")

 LcTotalGeral = Right("00000000" & LcTotalGeral, 8)
LcRegistro = LcRegistro & LcTotalGeral
Print #FnunBoleto, LcRegistro
Rs.Close
RsCliente.Close
RsDados.Close

Set Rs = Nothing
Set RsCliente = Nothing
Set RsDados = Nothing
Set db = Nothing
'set conexaoado = nothing
Close #FnunBoleto
'MsgBox "Total de Registros " & LcTotalRegistros
DisqueteReceita.Caption = LcCap
Screen.MousePointer = 0

Exit Function

errgera:
Screen.MousePointer = 0
If LcAbriu Then Close #FnunBoleto
MsgBox err.Description & err.Number
Resume 0

End Function
Function VerificaDisquete(drive As String)
On Error GoTo errveri

x = Dir(drive & "lidis.txt", vbArchive)
x = MsgBox("Todos os dados do Disquete seram Apagados, Confirma ?", vbExclamation + vbYesNo, "Aviso")
If x = 7 Then GoTo Saida
Kill drive & "*.*"


Saida:

Exit Function
errveri:
'Stop
'If err.Number = 52 Then
   x = MsgBox("Coloque o Disquete no drive <<A>> ou Remova a Proteção contra gravação." & Chr(13) & "clique Ok.", vbInformation + vbOKCancel, "Disquete não Encontrado.")
   If x = 2 Then Exit Function
   Resume 0
'End If
'MsgBox err.Description & "  " & err.Number

End Function
