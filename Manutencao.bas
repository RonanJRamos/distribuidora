Attribute VB_Name = "Manutencao"
Function AcertaEstoqueNovo()
Dim RsGalpao As ADODB.Recordset
Dim RsProduto As ADODB.Recordset

'abreconexao

Set RsGalpao = AbreRecordset("Select * from alid013 order by item", True)
Set RsProduto = AbreRecordset("Select * from Produtos order by codigo")
Do Until RsProduto.EOF
    Teste.Caption = "Produto:" & RsProduto!codigo & " " & RsProduto!Nome
    DoEvents
    RsGalpao.Filter = "item='" & Right("00000" & RsProduto!codigo, 5) & "'"
    Do Until RsGalpao.EOF
        Select Case RsGalpao!almox
            Case Is = "CALIFORNIA"
                RsProduto!california = (RsProduto!QtdMedida * RsGalpao!Estoque) + RsGalpao!quantUnidade
            Case Is = "SANTA MARIA"
                RsProduto!santa1 = (RsProduto!QtdMedida * RsGalpao!Estoque) + RsGalpao!quantUnidade
            Case Is = "SANTA MARIA 2"
                RsProduto!Santa2 = (RsProduto!QtdMedida * RsGalpao!Estoque) + RsGalpao!quantUnidade
        End Select
        
        RsGalpao.MoveNext
    Loop
    RsProduto!QuantEstoque = RsProduto!santa1 + RsProduto!Santa2 + RsProduto!california
    RsProduto.Update
    RsProduto.MoveNext
Loop
Teste.Caption = "Testes"
MsgBox "Terminei"
End Function
