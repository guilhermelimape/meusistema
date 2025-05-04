<%@ Language = VBScript %>
<!--#include file="includes/conexao.asp"-->
<%
Dim rs, sql, filtro
filtro = Trim(Request.QueryString("nome"))

' Tenta recuperar do cookie se não veio nada na URL
'if filtro = "" then
'    filtro = Trim(Request.Cookies("filtro_nome"))
'end if

' Salva no cookie se veio via URL
'if filtro <> "" then
'    Response.Cookies("filtro_nome") = filtro
'    Response.Cookies("filtro_nome").Expires = Date + 30 
'end if

sql = "select id, nome, email, idade, data_cadastro from msUsuarios"

if filtro <> "" then
    filtro = Replace(filtro, "'", "''") ' proteção básica
    sql = sql & " where nome like '%" & filtro & "%'"
end if

'sql = sql & " ORDER BY data_cadastro DESC"

Set rs = conn.Execute(sql)
%>

<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Lista de Usuários</title>
    </head>
<body>
    <form method="get" action="listar.asp">
        <label>Filtrar por nome:</label>
        <input type="text" name="nome" value="<%=Request.QueryString("nome")%>"> 
        <!--<input type="text" name="nome" value="<%'=filtro%>"> -->
        <input type="submit" value="Buscar">
    </form>
    <br>

    <h1>Usuários Cadastrados</h1>

    <table border="1" cellpadding="5">
        <tr>
            <th>ID</th>
            <th>Nome</th>
            <th>E-mail</th>
            <th>Idade</th>
            <th>Data de Cadastro</th>
            <th>Ações</th>
        </tr>

        <% Do until rs.eof %>
            <tr>
                <td><%=rs("id")%></td>
                <td><%=rs("nome")%></td>
                <td><%=rs("email")%></td>
                <td><%=rs("idade")%></td>
                <td><%=rs("data_cadastro")%></td>
                <td>
                    <a href="editar.asp?id=<%=rs("id")%>">✏️ Editar</a> |
                    <a href="excluir.asp?id=<%=rs("id")%>" onclick="return confirm('Tem certeza que deseja excluir?')">🗑 Excluir</a>
                </td>
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