<%@ Language=VBScript %>
<!--#include file="includes/conexao.asp"-->
<%

Dim usuario, senha, sql, rs
usuario = Trim(Request.Form("usuario"))
senha = Trim(Request.Form("senha"))


if usuario = "" or  senha = "" then
    response.write "Preencha todos os campos."
    response.end
end if

sql = "select * from msLogins where usuario = '" & usuario & "' and senha = '" & senha & "'"
set rs = conn.Execute(sql)

if not rs.eof then  
    session("usuario") = usuario
    response.redirect "listar.asp"
else
    response.write "Usuário ou senha inválidos."
end if

rs.Close
conn.Close
%>