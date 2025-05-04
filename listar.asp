<%@ Language = VBScript %>
<!--#include file="includes/conexao.asp"-->
<%
Dim re, sql
sql = "select id, nome, email, idade, data_cadastro from msUsuarios order by data_cadastro"
Set rs = conn.Execute(sql)
%>

<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Lista de Usuários</title>
    </head>
<body>
    <h1>Usuários Cadastrados</h1>

    <table border="1" cellpadding="5">
        <tr>
            <th>ID</th>
            <th>Nome</th>
            <th>E-mail</th>
            <th>Idade</th>
            <th>Data de Cadastro</th>
        </tr>

        <% Do until rs.eof %>
            <tr>
                <td><%=rs("id")%></td>
                <td><%=rs("nome")%></td>
                <td><%=rs("email")%></td>
                <td><%=rs("idade")%></td>
                <td><%=rs("data_cadastro")%></td>
            </tr>
            <%
            rs.MoveNext
        Loop
        %>

    </table>
    <p><a href="formulario.asp">← Voltar para o formulário</a></p>
</body>
</html>

<%
rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing
%>