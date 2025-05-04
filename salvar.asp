<%@ Language=VBScript %>
<!--#include file="includes/conexao.asp"-->
<%
Dim nome, cmd, email, idade, erro
nome = Trim(Request.Form("nome"))
email = Trim(Request.Form("email"))
idade = Trim(Request.Form("idade"))

erro = ""
' Validações
If nome = "" Then
    erro = erro & "<li>O campo <strong>Nome</strong> é obrigatório.</li>"
End If

if email = "" Then
    erro = erro & "<li>O campo <strong>Email</strong> é obrigatório.</li>"
elseif InStr(email, "@") = 0 or InStr(email, ".") = 0 Then
    erro = erro & "<li>O <strong>Email</strong> parece inválido.</li>"
end if 

if Not IsNumeric(idade) or CInt(idade) < 0 Then
    erro = erro & "<li>Informe uma <strong>idade válida</strong>.<li>"
end if 

If erro = "" then
    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandText = "insert into msUsuarios (nome, email, idade) values (?, ?, ?)"
    cmd.CommandType = 1 ' adCmdText

    cmd.Parameters.Append cmd.CreateParameter("@nome", 200, 1, 100, nome) 'adVarChar 
    cmd.Parameters.Append cmd.CreateParameter("@email", 200, 1, 100, email) 'adVarChar
    cmd.Parameters.Append cmd.CreateParameter("@idade", 3, 1, , CInt(idade)) 'adInteger

    cmd.Execute
    Set cmd = Nothing
    
    Response.Redirect "formulario.asp"
Else
%>
    <html><head><meta charset="UTF-8"></head><body>
    <h2>Erros encontrados:</h2>
    <ul style="color:red;">
        <%=erro%>
    </ul>
    <p><a href="javascript:history.back()">← Voltar e corrigir</a></p>
    </body></html>
<%
End If

conn.Close
Set conn = Nothing
%>
