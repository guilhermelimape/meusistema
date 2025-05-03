<%@ Language=VBScript %>
<!--#include file="includes/conexao.asp"-->
<%
Dim nome
nome = Trim(Request.Form("nome"))

If nome <> "" Then
    nome = Replace(nome, "'", "''") ' prevenção básica contra SQL Injection
    sql = "INSERT INTO msUsuarios (nome) VALUES ('" & nome & "')"
    conn.Execute sql
    Response.Redirect "formulario.asp"
Else
    Response.Write "Nome não pode ser vazio."
End If

conn.Close
Set conn = Nothing
%>
