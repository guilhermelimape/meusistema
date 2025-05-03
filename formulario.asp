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
    <h1>Cadastro de Nome</h1>

    <form method="post" action="salvar.asp">
        <label for="nome">Digite seu nome:</label>
        <input type="text" id="nome" name="nome" required>
        <input type="submit" value="Enviar">
    </form>

    <% if nome <> "" then %>
        <p>Olá, <strong><%=nome%></strong>! Seu nome foi recebido com sucesso.</p>
    <% end if %>
</body>
</html>