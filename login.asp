<%@ Language=VBScript %>
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Login</title>
</head>
<body>
  <h2>Entrar no sistema</h2>
    <form method="post" action="validar.asp">
        <label>Usuário:</label><br>
        <input type="text" name="usuario"><br><br>

        <label>Senha:</label><br>
        <input type="password" name="senha"><br><br>

        <input type="submit" value="Entrar">
  </form>
</body>
</html>