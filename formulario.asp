<%@ Language=VBScript %>

<%
' Recupera o valor enviado pelo formulário
Dim nome 
nome = Request.Form("nome")
%>

<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Formulário Simples</title>
</head>
<body>
    <p>Bem-vindo, <strong><%=Session("usuario")%></strong> | <a href="logout.asp">Sair</a></p>

    <h1>Cadastro de Usuário</h1>

    <form method="post" action="salvar.asp">
        <label for="nome">Nome</label><br>
        <input type="text" id="nome" name="nome"><br><br>

        <label>Email:</label><br>
        <input type="email" id="email" name="email"><br><br>

        <label>Idade:</label><br>
        <input type="number" id="idade" name="idade"><br><br>

        <input type="submit" value="Enviar">
    </form>

    <% if nome <> "" then %>
        <p>Olá, <strong><%=nome%></strong> cadastrado com sucesso.</p>
    <% end if %>

    <p><a href="listar.asp">→ Ver lista de usuários</a></p>
</body>
</html>