VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Utilitario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private ErrosP As String


Public Property Get erroS() As Variant
    
        erroS = ErrosP
    
End Property

Sub VerificaDigitacaoSubstituicao(DataInicio As Date, DataFinal As Date)
Dim Rs As ADODB.Recordset
Dim StrSql As String
ErrosP = ""
StrSql = "Select * from entradanf where DATA Between #" & Format(DataInicio, "mm/dd/yy") & "# And #" & Format(DataFinal, "mm/dd/yy") & "# and (BaseIcmsSubst>0 or IcmsSubst>0)"
Set Rs = AbreRecordset(StrSql, True)
'MsgBox DEscricaoErro
Debug.Print StrSql
'==> Recupera os dados e manda para ser gravado na matriz
Do Until Rs.EOF
   If Rs!BaseIcmsSubst = 0 Then
       ErrosP = ErrosP & "A Base de calculo do ICMS Subst. da NF:" & Rs!NF & " Esta zero." & Chr(13)
   End If
   If Rs!IcmsSubst = 0 Then
       ErrosP = ErrosP & "A Valor do ICMS Subst. da NF:" & Rs!NF & " Esta zero." & Chr(13)
   End If

   Rs.MoveNext
Loop
Rs.Close
Set Rs = Nothing
End Sub


Sub VerificaCnpj(Nome As String, Codigo As Long, Cliente As Boolean, Optional Cliente_Origem As String = "BHFORTE")
On Error Resume Next
Dim Rs As ADODB.Recordset
Dim StrSql As String
Dim Insc As String
Dim Cnpj As String
Dim db As Database
Dim Rsc As Recordset

If Cliente Then
  If Cliente_Origem = "BHFORTE" Then
    StrSql = "Select * from alid001 where codigo=" & Codigo
    Set Rs = AbreRecordset(StrSql, True)
    If Not Rs.EOF Then
        Cnpj = Replace(Rs!CGC, ".", "")
        Cnpj = Replace(Cnpj, ",", "")
        Cnpj = Replace(Cnpj, "-", "")
        Cnpj = Replace(Cnpj, "/", "")
        Cnpj = Replace(Cnpj, "\", "")
        Cnpj = Replace(Cnpj, " ", "")
        If Len(Cnpj) = 0 Then
           VerificaCPF Nome, Codigo, Cliente_Origem
           Exit Sub
        End If
        If Not Calc_CNPJ(Cnpj) Then
           ErrosP = ErrosP & " O CNPJ de " & Codigo & " - " & Nome & " é invalido." & Chr(13)
        End If
        Insc = Rs!INSCEST & ""
        Insc = Replace(Insc, ".", "")
        Insc = Replace(Insc, "-", "")
        Insc = Replace(Insc, "/", "")
        Insc = Replace(Insc, "\", "")
        Insc = Replace(Insc, " ", "")
        If Consiste(Insc, Rs!Estado) <> 0 Then
           ErrosP = ErrosP & " A Inscrição de " & Codigo & " - " & Nome & " é Inválida." & Chr(13)
        End If
    End If
    Set Rs = Nothing
  Else
     Set db = OpenDatabase(GLBase)
     StrSql = "Select * from alid001 where codigo='" & Right("00000" & Codigo, 5) & "'"
     Set Rsc = db.OpenRecordset(StrSql)
    If Not Rs.EOF Then
        Cnpj = Replace(Rsc!CGC, ".", "")
        Cnpj = Replace(Cnpj, ",", "")
        Cnpj = Replace(Cnpj, "-", "")
        Cnpj = Replace(Cnpj, "/", "")
        Cnpj = Replace(Cnpj, "\", "")
        Cnpj = Replace(Cnpj, " ", "")
        Cnpj = Replace(Cnpj, ",", "")
        If Len(Cnpj) = 0 Then
           VerificaCPF Nome, Codigo, Cliente_Origem
           Exit Sub
        End If
        If Not Calc_CNPJ(Cnpj) Then
           ErrosP = ErrosP & " O CNPJ de " & Codigo & " - " & Nome & " é invalido." & Chr(13)
        End If
        Insc = Rsc!INSCEST & ""
        Insc = Replace(Insc, ".", "")
        Insc = Replace(Insc, "-", "")
        Insc = Replace(Insc, "/", "")
        Insc = Replace(Insc, "\", "")
        Insc = Replace(Insc, ",", "")
        Insc = Replace(Insc, " ", "")
        If Consiste(Insc, Rsc!Estado) <> 0 Then
           ErrosP = ErrosP & " A Inscrição de " & Codigo & " - " & Nome & " é Inválida." & Chr(13)
        End If
    End If
    Set Rs = Nothing
  End If
Else
   
    If Cliente_Origem = "BHFORTE" Then
    StrSql = "Select * from Fornecedor where codigo=" & Codigo
    Set Rs = AbreRecordset(StrSql, True)
    If Not Rs.EOF Then
        Cnpj = Replace(Rs!CGC, ".", "")
        Cnpj = Replace(Cnpj, ",", "")
        Cnpj = Replace(Cnpj, "-", "")
        Cnpj = Replace(Cnpj, "/", "")
        Cnpj = Replace(Cnpj, "\", "")
        Cnpj = Replace(Cnpj, " ", "")
        If Len(Cnpj) = 0 Then
           VerificaCPF Nome, Codigo, Cliente_Origem
           Exit Sub
        End If
        If Not Calc_CNPJ(Cnpj) Then
           ErrosP = ErrosP & " O CNPJ de " & Codigo & " - " & Nome & " é invalido." & Chr(13)
        End If
        Insc = Rs!INSCEST & ""
        Insc = Replace(Insc, ".", "")
        Insc = Replace(Insc, "-", "")
        Insc = Replace(Insc, "/", "")
        Insc = Replace(Insc, "\", "")
        If Consiste(Insc, Rs!Estado) <> 0 Then
           ErrosP = ErrosP & " A Inscrição de " & Codigo & " - " & Nome & " é Inválida." & Chr(13)
        End If
    End If
    Set Rs = Nothing
  Else
     Set db = OpenDatabase(GLBase)
     StrSql = "Select * from alid002 where codigo='" & Right("00000" & Codigo, 5)
     Set Rsc = db.OpenRecordset(StrSql)
    If Not Rs.EOF Then
        Cnpj = Replace(Rsc!CGC, ".", "")
        Cnpj = Replace(Cnpj, ",", "")
        Cnpj = Replace(Cnpj, "-", "")
        Cnpj = Replace(Cnpj, "/", "")
        Cnpj = Replace(Cnpj, "\", "")
        Cnpj = Replace(Cnpj, " ", "")
        If Len(Cnpj) = 0 Then
           VerificaCPF Nome, Codigo, Cliente_Origem
           Exit Sub
        End If
        If Not Calc_CNPJ(Cnpj) Then
           ErrosP = ErrosP & " O CNPJ de " & Codigo & " - " & Nome & " é invalido." & Chr(13)
        End If
        Insc = Rsc!INSCEST & ""
        Insc = Replace(Insc, ".", "")
        Insc = Replace(Insc, "-", "")
        Insc = Replace(Insc, "/", "")
        Insc = Replace(Insc, "\", "")
        Insc = Replace(Insc, " ", "")
        If Consiste(Insc, Rsc!Estado) <> 0 Then
           ErrosP = ErrosP & " A Inscrição de " & Codigo & " - " & Nome & " é Inválida." & Chr(13)
        End If
    End If
    Set Rs = Nothing
  End If
End If



End Sub
Public Sub VerificaCPF(Nome As String, Codigo As Long, Optional Cliente_Origem As String = "BHFORTE")
Dim Rs As ADODB.Recordset
Dim StrSql As String
Dim Insc As String
Dim cpf As String
Dim db As Database
Dim Rsc As Recordset
If Cliente_Origem = "BHFORTE" Then
   StrSql = "Select * from alid001 where codigo=" & Codigo
   Set Rs = AbreRecordset(StrSql, True)
    If Not Rs.EOF Then
        cpf = Replace(Rs!cpf, ".", "")
        cpf = Replace(cpf, ",", "")
        cpf = Replace(cpf, "-", "")
        cpf = Replace(cpf, "/", "")
        cpf = Replace(cpf, "\", "")
        cpf = Replace(cpf, " ", "")
        If Not Calc_CPF(cpf) Then
           ErrosP = ErrosP & " O CPF de " & Codigo & " - " & Nome & " é invalido." & Chr(13)
        End If
        Insc = Rs!INSCEST & ""
        Insc = Replace(Insc, ".", "")
        Insc = Replace(Insc, "-", "")
        Insc = Replace(Insc, "/", "")
        Insc = Replace(Insc, "\", "")
        If Consiste(Insc, Rs!Estado) <> 0 Then
           ErrosP = ErrosP & " A Inscrição de " & Codigo & " - " & Nome & " é Inválida." & Chr(13)
        End If
    End If
    Set Rs = Nothing
Else
   Set db = OpenDatabase(GLBase)
   
   StrSql = "Select * from alid001 where codigo='" & Right("00000" & Codigo, 5) & "'"
   Set Rsc = db.OpenRecordset(StrSql)
    If Not Rsc.EOF Then
        cpf = Replace(Rsc!cpf, ".", "")
        cpf = Replace(cpf, ",", "")
        cpf = Replace(cpf, "-", "")
        cpf = Replace(cpf, "/", "")
        cpf = Replace(cpf, "\", "")
        cpf = Replace(cpf, " ", "")
        If Not Calc_CPF(cpf) Then
           ErrosP = ErrosP & " O CPF de " & Codigo & " - " & Nome & " é invalido." & Chr(13)
        End If
        Insc = Rsc!INSCEST & ""
        Insc = Replace(Insc, ".", "")
        Insc = Replace(Insc, "-", "")
        Insc = Replace(Insc, "/", "")
        Insc = Replace(Insc, "\", "")
        If Consiste(Insc, Rsc!Estado) <> 0 Then
           ErrosP = ErrosP & " A Inscrição de " & Codigo & " - " & Nome & " é Inválida." & Chr(13)
        End If
    End If
    Set Rsc = Nothing
End If

End Sub
Sub VerificaValorProduto(ValorProduto As Double, CodigoNota As String, CodigoFornecedor As Long, Optional CodigoDaNota As Long = 0)
Dim RsItem As ADODB.Recordset
Dim StrSql As String
Dim LcValor As Double
If CodigoDaNota = 0 Then
    StrSql = "SELECT Sum(ItensEntradaNf.Valor_Total_Compra) AS SomaDeValor_Total_Compra " & _
         "From ItensEntradaNf WHERE (((ItensEntradaNf.NUMNF)='" & CodigoNota & "') AND ((ItensEntradaNf.codigofor)=" & CodigoFornecedor & "));"
Else
    StrSql = "SELECT Sum(ItensEntradaNf.Valor_Total_Compra) AS SomaDeValor_Total_Compra " & _
         "From ItensEntradaNf WHERE CodigoNota=" & CodigoDaNota & ";"

End If
'If CodigoNota = "28" Then MsgBox StrSql
Set RsItem = AbreRecordset(StrSql, True)

Debug.Print StrSql

Do Until RsItem.EOF
   LcValor = IIf(Not IsNull(RsItem!SomaDeValor_Total_Compra), RsItem!SomaDeValor_Total_Compra, 0)
   RsItem.MoveNext
Loop
Set RsItem = Nothing

If CCur(LcValor) <> CCur(ValorProduto) Then
   ErrosP = ErrosP & " Nota de entrada Nº:" & CodigoNota & " Valor total de produto não confere com o valor da nota."
End If
End Sub

Public Function AcertaNumero(LcNumero As String, LcTamaDecimal As Long) As String
Dim LcTa, a As Long
Dim LcDecimal, LcInteiro, LCLEtra, LcZeros As String
Dim LcAchou As Integer
0 LcNumero = CDbl(LcNumero)
LcNumero = CStr(LcNumero)
If Len(LcNumero) = 0 Then AcertaNumero = 0: Exit Function

If LcTamaDecimal = 0 Then LcTamaDecimal = 2
LcTa = Len(LcNumero)
For a = 1 To LcTa
    LCLEtra = Mid$(LcNumero, a, 1)
    If LCLEtra = "," Or LCLEtra = "." Then
       LcAchou = True
    End If
    If LcAchou Then
       If LCLEtra <> "," Or LCLEtra = "." Then
          LcDecimal = LcDecimal & LCLEtra
       End If
    Else
       LcInteiro = LcInteiro & LCLEtra
    End If
Next
If Len(LcInteiro) = 0 Then LcInteiro = "0"
For a = 1 To LcTamaDecimal
    LcZeros = LcZeros & "0"
Next
LcDecimal = Left(LcDecimal & LcZeros, LcTamaDecimal)
AcertaNumero = LcInteiro & "," & LcDecimal
End Function

Private Function Calc_CNPJ(CGC As String) As Boolean
  Dim Retorno, a, j, i, d1, d2
  'TiraMascara CGC
  
  CGC = Replace(CGC, ".", "")
  CGC = Replace(CGC, ",", "")
  CGC = Replace(CGC, "-", "")
  CGC = Replace(CGC, "/", "")
  CGC = Replace(CGC, "\", "")
  CGC = Replace(CGC, " ", "")
  'Debug.Print CGC
  If Len(CGC) = 8 And Val(CGC) > 0 Then
     a = 0
     j = 0
     d1 = 0
     For i = 1 To 7
         a = Val(Mid(CGC, i, 1))
         If (i Mod 2) <> 0 Then
            a = a * 2
         End If
         If a > 9 Then
            j = j + Int(a / 10) + (a Mod 10)
         Else
            j = j + a
         End If
     Next i
     d1 = IIf((j Mod 10) <> 0, 10 - (j Mod 10), 0)
     If d1 = Val(Mid(CGC, 8, 1)) Then
        ValidaCGC = True
     Else
        ValidaCGC = False
        'MsgBox "CNPJ inválido!,Verifique", vbCritical, "Valida CGC"
     End If
  Else
     If Len(CGC) = 14 And Val(CGC) > 0 Then
        a = 0
        i = 0
        d1 = 0
        d2 = 0
        j = 5
        For i = 1 To 12 Step 1
            a = a + (Val(Mid(CGC, i, 1)) * j)
            j = IIf(j > 2, j - 1, 9)
        Next i
        a = a Mod 11
        d1 = IIf(a > 1, 11 - a, 0)
        a = 0
        i = 0
        j = 6
        For i = 1 To 13 Step 1
            a = a + (Val(Mid(CGC, i, 1)) * j)
            j = IIf(j > 2, j - 1, 9)
        Next i
        a = a Mod 11
        d2 = IIf(a > 1, 11 - a, 0)
        If (d1 = Val(Mid(CGC, 13, 1)) And d2 = Val(Mid(CGC, 14, 1))) Then
           ValidaCGC = True
        Else
           ValidaCGC = False
          ' MsgBox "CNPJ inválido!, Verifique", vbCritical, "Aviso do Sistema"
        End If
     Else
        ValidaCGC = False
        'MsgBox "CNPJ inválido!,Verifique", vbCritical, "Aviso do Sistema"
     End If
  End If
  Calc_CNPJ = ValidaCGC
End Function
Private Function Calc_CPF(valor As String) As Boolean
' Inicializa variaveis
Dim dig1 As Integer
Dim dig2 As Integer
Dim Mult1 As Integer
Dim Mult2 As Integer
Dim x As Integer
Dim a As Integer
Dim LcCpf As String
Mult1 = 10
Mult2 = 11
For a = 1 To Len(valor)
    If IsNumeric(Mid(valor, a, 1)) Then
       LcCpf = LcCpf & Mid(valor, a, 1)
    End If
Next
valor = LcCpf
For x = 1 To 9
dig1 = dig1 + (Val(Mid$(valor, x, 1)) * Mult1)
Mult1 = Mult1 - 1
Next

For x = 1 To 10
dig2 = dig2 + (Val(Mid$(valor, x, 1)) * Mult2)
Mult2 = Mult2 - 1
Next

dig1 = (dig1 * 10) Mod 11
dig2 = (dig2 * 10) Mod 11
If dig1 = 10 Then dig1 = 0
If dig2 = 10 Then dig2 = 0

Calc_CPF = False

If Val(Mid$(valor, 10, 1)) = dig1 And Val(Mid$(valor, 11, 1)) = dig2 Then Calc_CPF = True
'If Val(Mid$(VALOR, 11, 1)) <> dig2 Then Calc_CPF = True
End Function
Private Function Consiste(Insc As String, UF As String) As Integer
  Consiste = ConsisteInscricaoEstadual(Insc, UF)
End Function

