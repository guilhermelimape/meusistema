<%@ Language=VBScript %>
<!--#include file="includes/conexao.asp"-->
<%
' Recupera o ID enviado na URL (ex: editar.asp?id=3)
Dim id, rs, sql
id = Trim(Request.QueryString("id"))

If Not IsNumeric(id) Then
  Response.Write "ID inválido."
  Response.End
End If

' Busca o registro no banco com o ID
sql = "SELECT * FROM msUsuarios WHERE id = " & id
Set rs = conn.Execute(sql)

If rs.EOF Then
  Response.Write "Usuário não encontrado."
  Response.End
End If

' Variáveis com os dados atuais
Dim nome, email, idade
nome = rs("nome")
email = rs("email")
idade = rs("idade")

rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing
%>

<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Editar Usuário</title>
</head>
<body>
  <h1>Editar Usuário</h1>

  <form method="post" action="atualizar.asp">
    <input type="hidden" name="id" value="<%=id%>">
    
    <label>Nome:</label><br>
    <input type="text" name="nome" value="<%=nome%>" required><br><br>

    <label>Email:</label><br>
    <input type="email" name="email" value="<%=email%>" required><br><br>

    <label>Idade:</label><br>
    <input type="number" name="idade" value="<%=idade%>" required><br><br>

    <input type="submit" value="Salvar alterações">
    <a href="listar.asp">Cancelar</a>
  </form>
</body>
</html>
