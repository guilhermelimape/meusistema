<%@ Language=VBScript %>
<!--#include file="includes/conexao.asp"-->
<%
Dim nome, cmd
nome = Trim(Request.Form("nome"))

If nome <> "" Then

    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandText = "insert into msUsuarios (nome) values (?)"
    cmd.CommandType = 1 ' adCmdText

    cmd.Parameters.Append cmd.CreateParameter("@nome", 200, 1, 100, nome)

    cmd.Execute
    Set cmd = Nothing
    
    Response.Redirect "formulario.asp"
Else
    Response.Write "Nome não pode ser vazio."
End If

conn.Close
Set conn = Nothing
%>
