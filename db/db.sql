DROP TABLE IF EXISTS msUsuarios;
DROP TABLE IF EXISTS msLogins;

CREATE TABLE msUsuarios (
  id INT IDENTITY PRIMARY KEY,
  nome VARCHAR(100),
  email VARCHAR(100),
  idade INT,
  data_cadastro DATETIME DEFAULT GETDATE()
);

CREATE TABLE msLogins (
  id INT IDENTITY PRIMARY KEY,
  usuario VARCHAR(50),
  senha VARCHAR(50)
);

INSERT INTO msLogins (usuario, senha) VALUES ('admin', '123');
