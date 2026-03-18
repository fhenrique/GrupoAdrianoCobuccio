USE master;
GO

--Database
IF EXISTS (SELECT name FROM sys.databases WHERE name = N'DBEmpresaxyz')
    DROP DATABASE DBAdrianoCobuccio;
GO

CREATE DATABASE DBEmpresaxyz;
GO

USE DBEmpresaxyz;
GO

-- Tabela Transacao
CREATE TABLE Transacao (
    Id_Transacao INT PRIMARY KEY IDENTITY(1,1), 
    Numero_Cartao VARCHAR(16) NOT NULL,
    Valor_Transacao DECIMAL(15, 2),
    Data_Transacao DATE DEFAULT GETDATE(),
    Descricao VARCHAR(150)
);
GO

-- Tabela Clientes
CREATE TABLE Clientes (
    Id_Cliente INT PRIMARY KEY IDENTITY(1,1), 
    Nome_Cliente VARCHAR(50) NOT NULL,
	Numero_Cartao VARCHAR(16) NOT NULL
);
GO

--Procedure
CREATE PROCEDURE sp_TotalTransacoesPorCartao
    @Data_Inicial DATE,
    @Data_Final DATE
AS
BEGIN
    SELECT 
        Numero_Cartao,
        COUNT(*) AS Total_Transacoes,
        SUM(Valor_Transacao) AS Quantidade_Transacoes
    FROM 
        Transacao
    WHERE 
        Data_Transacao BETWEEN @Data_Inicial AND @Data_Final
    GROUP BY 
        Numero_Cartao;
END;
GO


--Função
CREATE FUNCTION fn_CategoriaTransacao
(
    @Valor_Transacao DECIMAL(15, 2)
)
RETURNS VARCHAR(10)
AS
BEGIN
    DECLARE @Categoria VARCHAR(10);

    IF @Valor_Transacao > 1000
        SET @Categoria = 'alta';
    ELSE IF @Valor_Transacao BETWEEN 500 AND 1000
        SET @Categoria = 'média';
    ELSE
        SET @Categoria = 'baixa';

    RETURN @Categoria;
END;
GO

-- View
CREATE VIEW vw_TransacoesClientes
AS
SELECT 
    cli.Nome_Cliente,
    tra.Numero_Cartao,
    tra.Valor_Transacao,
    tra.Data_Transacao,
    dbo.fn_CategoriaTransacao(tra.Valor_Transacao) AS Categoria
FROM 
    Transacao tra
INNER JOIN Clientes cli ON (tra.Numero_Cartao = cli.Numero_Cartao);
GO