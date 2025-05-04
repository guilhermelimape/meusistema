<%@ Language=VBScript %>
<!--#include file="includes/conexao.asp"-->
<%
Dim id, sql
id = Trim(Request.QueryString("id"))

If Not IsNumeric(id) Then
    Response.Write "ID inválido."
    Response.End
End If

sql = "DELETE FROM msUsuarios WHERE id = " & id

conn.Execute sql

conn.Close
Set conn = Nothing

Response.Redirect "listar.asp"

%>