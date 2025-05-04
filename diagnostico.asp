<%@ Language=VBScript %>
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Diagnóstico do Ambiente</title>
</head>
<body>
    <h1>Informações do Visitante</h1>
    <ul>
        <li><strong>IP do usuário:<strong> <%= Request.ServerVariables("REMOTE_ADDR") %></li>
        <li><strong>Navegador:</strong> <%= Request.ServerVariables("HTTP_USER_AGENT")%></li>
        <li><strong>Endereço da página:</strong> <%= Request.ServerVariables("URL") %></li>
        <li><strong>Referência (HTTP_REFERER):</strong> <%= Request.ServerVariables("HTTP_REFERER") %></li>
        <li><strong>Servidor (SERVER_NAME):</strong> <%= Request.ServerVariables("SERVER_NAME") %></li>
        <li><strong>Método da requisição:</strong> <%= Request.ServerVariables("REQUEST_METHOD") %></li>
    </ul>

    <p><a href="formulario.asp">← Voltar para o formulário</a></p>
</body>
</html>