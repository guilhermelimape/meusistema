<%@ Language=VBScript %>
<!--#include file="includes/conexao.asp"-->
<%
Dim id, nome, email, idade, sql

id = Trim(Request.Form("id"))
nome = Replace(Trim(Request.Form("nome")), "'", "''")
email = Replace(Trim(Request.Form("email")), "'", "''")
idade = Trim(Request.Form("idade"))

If Not IsNumeric(id) Then
    Response.Write "ID inválido."
    Response.End
End If

If nome = "" Or email = "" Or Not IsNumeric(idade) Then
    Response.Write "Dados inválidos. Verifique os campos."
    Response.End
End If

sql = "UPDATE msusuarios SET nome = '" & nome & "', email = '" & email & "', idade = " & idade & " WHERE id = " & id

conn.Execute sql

conn.Close
Set conn = Nothing

Response.Redirect "listar.asp"
%>