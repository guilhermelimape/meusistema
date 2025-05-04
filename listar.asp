<%@ Language = VBScript %>
<!--#include file="includes/conexao.asp"-->
<%

If Session("usuario") = "" Then
  Response.Redirect "login.asp"
End If

Dim rs, sql, filtro, pagina, paginaAtual, registrosPorPagina

pagina = Request.QueryString("pagina")
if not IsNumeric(pagina) or pagina = "" then 
    paginaAtual = 1
else
    paginaAtual = CInt(pagina)
end if

registrosPorPagina = 5 ' Você pode ajustar

filtro = Trim(Request.QueryString("nome"))

sql = "select id, nome, email, idade, data_cadastro from msUsuarios"

if filtro <> "" then
    filtro = Replace(filtro, "'", "''") ' proteção básica
    sql = sql & " where nome like '%" & filtro & "%'"
end if
'sql = sql & " ORDER BY data_cadastro DESC"

Set rs = Server.CreateObject("ADODB.Recordset")
rs.PageSize = registrosPorPagina
rs.CacheSize = registrosPorPagina
rs.Open sql, conn, 1, 3 'adOpenKeyset, adLockOptimistic

if not rs.eof then
    rs.AbsolutePage = paginaAtual
end if

Dim totalPaginas
if rs.RecordCount > 0 then
    totalPaginas = rs.PageCount
else
    totalPaginas = 1
end if

'Set rs = conn.Execute(sql)
%>

<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Lista de Usuários</title>
    </head>
<body>
    <p>Bem-vindo, <strong><%=Session("usuario")%></strong> | <a href="logout.asp">Sair</a></p>

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

        <%
        Dim contador
        contador = 1
        Do while not rs.eof and contador <= registrosPorPagina
        %>
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
            contador = contador + 1
        Loop
        %>

    </table>
    <p>
    <%
    if paginaAtual > 1 then 
        response.write "<a href='listar.asp?pagina=" & (paginaAtual - 1)
        if filtro <> "" then response.write "&nome=" & Server.URLEncode(filtro)
        response.write "'>◀ Anterior</a> "
    end if

    for i = 1 to totalPaginas
        if i = paginaAtual then
            response.write " <strong>[" & i & "]</strong> "
        else    
            response.write " <a href='listar.asp?pagina=" & i
        If filtro <> "" Then Response.Write "&nome=" & Server.URLEncode(filtro)
            Response.Write "'>" & i & "</a> "
        End If
    Next
    
        If paginaAtual < totalPaginas Then
            Response.Write " <a href='listar.asp?pagina=" & (paginaAtual + 1)
        If filtro <> "" Then Response.Write "&nome=" & Server.URLEncode(filtro)
             Response.Write "'>Próxima ▶</a>"
        End If
    %>
    </p>
   

</body>
</html>

<%
rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing
%>