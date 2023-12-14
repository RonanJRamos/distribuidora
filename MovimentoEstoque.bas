Attribute VB_Name = "MovimentoEstoque"
Function BaixaPorNota(codigo As String, Quantidade As Double, com As Double, Unidade As String, CodUnid As String)
Dim RsH         As Recordset

Dim LcTotalUnit As Double
Dim LcTotalCx   As Double
Dim LcUnGalpao  As Double
Dim LcCxGalpao  As Double
Dim LcUndPr     As String
Dim LcQuantUn   As Double
Dim LcUnSanta   As Double
Dim LcCxSanta   As Double
Dim LcUnSanta1  As Double
Dim LcCxSanta1  As Double
Dim LcUnCali    As Double
Dim LcCxCali    As Double
Dim LcUnidBaixa As Double
Dim LcCaixaBaixa As Double '

Dim LcQuantUnBasica As Double
Dim LcQuantBasica As Double
Dim LcQuantcxVend As Double
Dim DATA As Date
Dim santa As Double
Dim santa1 As Double
Dim california As Double
Dim santau As Double
Dim santa1u As Double
Dim californau As Double
Dim LcBaixou As Boolean
AbreBase
Set RsM = Dbbase.OpenRecordset("select * from alid009 where cod='" & codigo & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsG = Dbbase.OpenRecordset("Select * from Alid013 where item='" & codigo & "' order by codigogalpao", dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsH = Dbbase.OpenRecordset("Select * from HistoricoProduto where produto='" & codigo & "' order by codigo", dbOpenDynaset, dbSeeChanges, dbOptimistic) '
'===> Vamos converter tudo para a unidade basica
LcQuantUnBasica = com * Quantidade
'===> Vamos Pegar a unidade Basica
LcQuantBasica = RsM!QTDUNIMED
LcQuantcxVend = Int(LcQuantUnBasica / LcQuantBasica)
lcquantunvend = ((LcQuantUnBasica / LcQuantBasica) - LcQuantcxVend) * LcQuantBasica
If lcquantunvend < 0 Then lcquantunvend = lcquantunvend * (-1)
If RsG.EOF Then
   RsG.AddNew
   RsG!Item = codigo
   RsG!almox = "SANTA MARIA"
   RsG!Estoque = 0
   RsG!Descricao = RsM!nome
   RsG!quantUnidade = 0
   RsG!CODIGOGALPAO = "02"
   RsG.Update
  
   RsG.AddNew
   RsG!Item = codigo
   RsG!almox = "SANTA MARIA 2"
   RsG!Estoque = 0
   RsG!Descricao = RsM!nome
   RsG!quantUnidade = 0
   RsG!CODIGOGALPAO = "02"
   RsG.Update
  
   RsG.AddNew
   RsG!Item = codigo
   RsG!almox = "CALIFORNIA"
   RsG!Estoque = 0
   RsG!Descricao = RsM!nome
   RsG!quantUnidade = 0
   RsG!CODIGOGALPAO = "02"
   RsG.Update
End If
RsG.MoveFirst
Do Until RsG.EOF
   '===> Verifica se o galpao tem a Quantidade para baixar o estoque
   If ((RsG!Estoque * LcQuantBasica) + RsG!quantUnidade) >= LcQuantUnBasica Then
       '===>Tem a Quantidade, entao vamos baixar
      '===> Vamos Gerar o Estoque
       Select Case RsG!almox
         Case "SANTA MARIA"
            santa = LcQuantcxVend
            santau = lcquantunvend
            santa1 = 0
            california = 0
            santa1u = 0
            californau = 0
         Case "SANTA MARIA 2"
            santa = 0
            santau = 0
            santa1 = LcQuantcxVend
            california = 0
            santa1u = lcquantunvend
            californau = 0
         Case "CALIFORNIA"
             santa = 0
            santau = 0
            santa1 = 0
            california = LcQuantcxVend
            santa1u = 0
            californau = lcquantunvend
       End Select
       DATA = CDate(GlFormA.Txt(12).Text)
       Descricao = RsM!nome
       Call GeraHistorico(codigo, Descricao, CStr(GlFormA.Txt(0).Text), "S", DATA, santa, santa1, california, santau, santa1u, californau)
       RsG.Edit
       RsG!Estoque = RsG!Estoque - LcQuantcxVend
       If RsG!quantUnidade < lcquantunvend Then
           RsG!quantUnidade = (RsG!quantUnidade + LcQuantBasica) - lcquantunvend
           RsG!Estoque = RsG!Estoque - 1
       Else
          RsG!quantUnidade = RsG!quantUnidade - lcquantunvend
       End If
       If RsG!quantUnidade < 0 Then RsG!quantUnidade = RsG!quantUnidade * (-1)
       RsG.Update
       LcBaixou = True
       Exit Do
   Else
      '===> Não tem a quantidade nescessaria
      '===> Vamos Ver se tem Alguma Coisa no galpao
      If ((RsG!Estoque * LcQuantBasica) + RsG!quantUnidade) > 0 Then
         '===> Tem Algo no Galpao, Entao vamos Tirar
                Select Case RsG!almox
         Case "SANTA MARIA"
            santa = RsG!Estoque
            santau = RsG!quantUnidade
            santa1 = 0
            california = 0
            santa1u = 0
            californiau = 0
         Case "SANTA MARIA 2"
            santa = 0
            santau = 0
            santa1 = RsG!Estoque
            california = 0
            santa1u = RsG!quantUnidade
            californiau = 0
         Case "CALIFORNIA"
             santa = 0
            santau = 0
            santa1 = 0
            california = RsG!Estoque
            santa1u = 0
            californiau = RsG!quantUnidade
      End Select
       DATA = CDate(GlFormA.Txt(12).Text)
       Descricao = RsM!nome
       Call GeraHistorico(codigo, Descricao, NF, "S", DATA, santa, santa1, california, santau, santa1u, californau) '
         LcQuantcxVend = LcQuantcxVend - RsG!Estoque
        lcquantunvend = lcquantunvend - RsG!quantUnidade
         RsG.Edit
         RsG!Estoque = 0
         RsG!quantUnidade = 0
         RsG.Update
      End If
   End If
   If (LcQuantcxVend + lcquantunvend) <= 0 Then Exit Do
   RsG.MoveNext
Loop
'If Not LcBaixou Then
'   If (lcquantunvend + LcQuantcxVend) > 0 Then
'      LcPes = "ALMOX='SANTA MARIA'"
'      RsG.FindFirst LcPes
'      If Not RsG.NoMatch Then
'         RsG.Edit
'         RsG!Estoque = RsG!Estoque - LcQuantcxVend
'         RsG!quantUnidade = RsG!quantUnidade - lcquantunvend
'         RsG.Update
'      End If
'   End If
'End If


'===>Vamos Corrigir o Estoque
If Not RsG.EOF Then RsG.MoveFirst
LcSaldocx = 0
LcSaldoUn = 0
Do Until RsG.EOF
   If Not IsNull(RsG!Estoque) Then LcSaldocx = LcSaldocx + RsG!Estoque
   If Not IsNull(RsG!quantUnidade) Then LcSaldoUn = LcSaldoUn + RsG!quantUnidade
   RsG.MoveNext
Loop
RsM.Edit
RsM!QuantEstoque = LcSaldocx
RsM!quantUnidade = LcSaldoUn
RsM.Update '
'RsM.Close
RsG.Close
Dbbase.Close
Set RsM = Nothing
Set RsG = Nothing
Set Dbbase = Nothing '

    
End Function
Function estornonotasaida(NF As String, codigo As String)
Dim RsH As Recordset
Dim Rsp As Recordset
Dim RsG As Recordset
Dim Rsc As Recordset
Dim LcQuant As Double
AbreBase
DoEvents
lcsq = "Select * from HistoricoProduto where produto='" & codigo & "' and nf='" & NF & "' and tipo='S'"
LcSql = "Select * from alid009 where cod='" & codigo & "'"
LcSql1 = "Select * from alid013 where item='" & codigo & "'"
lcsqlcl = "Select * from alid001"
Set RsH = Dbbase.OpenRecordset(lcsq, dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set Rsp = Dbbase.OpenRecordset(LcSql, dbOpenDynaset, dbSeeChanges, dbOptimistic)
Set RsG = Dbbase.OpenRecordset(LcSql1, dbOpenDynaset, dbSeeChanges, dbOptimistic)
'Set Rsc = Dbbase.OpenRecordset(lcsqlcl)
LcQuant = 0
If Not Rsp.EOF Then
   LcQuant = Rsp!QTDUNIMED
End If
If Not RsH.EOF Then
'===> Vamos verificar se a saida foi do santa maria
   If RsH!santa + RsH!unisanta > 0 Then
      '===> Tem entrada dentro do santa maria
      LcPes = "ALMOX='SANTA MARIA'"
      RsG.FindFirst LcPes
      If Not RsG.NoMatch Then
         RsG.Edit
         RsG!Estoque = RsG!Estoque + RsH!santa
         If (RsH!unisanta + RsG!quantUnidade) > LcQuant Then
            lcnsaldo = (RsH!unisanta + RsG!quantUnidade) - LcQuant
            RsG!quantUnidade = lcnsaldo
            RsG!Estoque = RsG!Estoque + 1
         Else
            RsG!quantUnidade = RsG!quantUnidade + RsH!unisanta
         End If
         RsG.Update
       End If
    End If
'===> Vamos verificar se a saida foi do santa maria 1
   If RsH!santa2 + RsH!unsanta1 > 0 Then
      '===> Tem entrada dentro do santa maria
      LcPes = "ALMOX='SANTA MARIA 2'"
      RsG.FindFirst LcPes
      If Not RsG.NoMatch Then
         RsG.Edit
         RsG!Estoque = RsG!Estoque + RsH!santa2
         If (RsH!unsanta1 + RsG!quantUnidade) > LcQuant Then
            lcnsaldo = (RsH!unsanta1 + RsG!quantUnidade) - LcQuant
            RsG!quantUnidade = lcnsaldo
            RsG!Estoque = RsG!Estoque + 1
         Else
            RsG!quantUnidade = RsG!quantUnidade + RsH!unsanta1
         End If
         RsG.Update
       End If
    End If
'===> Vamos verificar se a saida foi do santa california
   If RsH!california + RsH!Uncalifornia > 0 Then
      '===> Tem entrada dentro do santa maria
      LcPes = "ALMOX='CALIFORNIA'"
      RsG.FindFirst LcPes
      If Not RsG.NoMatch Then
         RsG.Edit
         RsG!Estoque = RsG!Estoque + RsH!california
         If (RsH!Uncalifornia + RsG!quantUnidade) > LcQuant Then
            lcnsaldo = (RsH!Uncalifornia + RsG!quantUnidade) - LcQuant
            RsG!quantUnidade = lcnsaldo
            RsG!Estoque = RsG!Estoque + 1
         Else
            RsG!quantUnidade = RsG!quantUnidade + RsH!Uncalifornia
         End If
         RsG.Update
       End If
    End If
End If
'===>Vamos atualizar o estoque
RsG.Close
Set RsG = Dbbase.OpenRecordset("select * from alid013 where item='" & codigo & "'", dbOpenDynaset, dbSeeChanges, dbOptimistic)
LcSaldocx = 0
LcSaldoUn = 0
Do Until RsG.EOF
   DoEvents
   If Not IsNull(RsG!Estoque) Then LcSaldocx = LcSaldocx + RsG!Estoque
   If Not IsNull(RsG!quantUnidade) Then LcSaldoUn = LcSaldoUn + RsG!quantUnidade
   RsG.MoveNext
Loop
Rsp.Edit
Rsp!QuantEstoque = LcSaldocx
Rsp!quantUnidade = LcSaldoUn
Rsp.Update

'Rsp.Close
'RsG.Close
'Rsp.Close
'Dbbase.Close

'Set Rsp = Nothing
'Set RsG = Nothing
'Set Rsp = Nothing
'Set Dbbase = Nothing

End Function
