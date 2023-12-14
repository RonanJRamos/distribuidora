INSERT INTO estoquefiscal ( Nome, CodigoProduto, Data )
SELECT produtos.NOME, produtos.codigo, 31/12/3 AS Expr1
FROM produtos;
