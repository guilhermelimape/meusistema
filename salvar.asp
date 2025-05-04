<%@ Language=VBScript %>
<!--#include file="includes/conexao.asp"-->
<%
Dim nome, cmd
nome = Trim(Request.Form("nome"))
email = Trim(Request.Form("email"))
idade = Trim(Request.Form("idade"))

If nome <> "" Then
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
    Response.Write "Nome não pode ser vazio."
End If

conn.Close
Set conn = Nothing
%>
