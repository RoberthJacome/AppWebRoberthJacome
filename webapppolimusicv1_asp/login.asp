<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<% Option Explicit %>
<%
Response.CodePage = 65001
Response.Charset = "utf-8"
%>
<!--#include file="includes/conexion.asp"-->
<%
' ---------- Login ----------
Dim usuario, clave, errorMsg, conn, cmd, rs, idUser, userName, userType

' Si ya hay sesión, directo a la lista
If Trim(Session("Username") & "") <> "" Then
    Response.Redirect "songs.asp"
End If

usuario  = ""
errorMsg = ""

If UCase(Request.ServerVariables("REQUEST_METHOD")) = "POST" Then

    usuario = Trim(Request.Form("usuario") & "")
    clave   = Request.Form("clave") & ""

    If usuario = "" Or clave = "" Then
        errorMsg = "Escribe tu usuario y tu contraseña."
    Else
        Set conn = AbrirConexion()

        ' Consulta parametrizada con ADODB.Command: nunca se concatena SQL
        Set cmd = Server.CreateObject("ADODB.Command")
        Set cmd.ActiveConnection = conn
        cmd.CommandType = adCmdText
        cmd.CommandText = "SELECT ID_USER, USERNAME, USER_TYPE FROM TBL_USER WHERE USERNAME = ? AND PASSWORD = ?"
        cmd.Parameters.Append cmd.CreateParameter("pUsuario", adVarWChar, adParamInput, 50, usuario)
        cmd.Parameters.Append cmd.CreateParameter("pClave", adVarWChar, adParamInput, 100, clave)

        Set rs = cmd.Execute

        If Not rs.EOF Then
            ' Credenciales correctas -> guardar sesión
            idUser   = rs("ID_USER")
            userName = rs("USERNAME") & ""
            userType = 0
            If Not IsNull(rs("USER_TYPE")) Then userType = CInt(rs("USER_TYPE"))

            rs.Close
            Set rs = Nothing
            conn.Close
            Set conn = Nothing

            Session("UserID")   = idUser
            Session("Username") = userName
            Session("UserType") = userType

            Response.Redirect "songs.asp"
        Else
            errorMsg = "Usuario o contraseña incorrectos."

            rs.Close
            Set rs = Nothing
            conn.Close
            Set conn = Nothing
        End If
    End If

End If
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Iniciar sesión · PoliMusic</title>
<link rel="stylesheet" href="css/estilos.css">
</head>
<body class="login-body">

    <main class="login-card">

        <div class="brand brand-lg">
            <span class="brand-mark">♪</span>
            <span class="brand-text">Poli<span>Music</span></span>
        </div>

        <h1>Bienvenido de nuevo</h1>
        <p class="login-sub">Inicia sesión para escuchar tu música.</p>

        <% If errorMsg <> "" Then %>
        <div class="alert" role="alert">
            <span class="alert-ico">!</span>
            <span><%= Server.HTMLEncode(errorMsg) %></span>
        </div>
        <% End If %>

        <form method="post" action="login.asp" autocomplete="off">
            <div class="field">
                <label for="usuario">Usuario</label>
                <input class="input" type="text" id="usuario" name="usuario"
                       value="<%= Server.HTMLEncode(usuario) %>"
                       placeholder="Tu usuario" maxlength="50" required autofocus>
            </div>

            <div class="field">
                <label for="clave">Contraseña</label>
                <input class="input" type="password" id="clave" name="clave"
                       placeholder="••••••••" maxlength="100" required>
            </div>

            <button type="submit" class="btn btn-primary btn-block">Entrar</button>
        </form>

        <p class="login-foot">PoliMusic · ASP clásico + SQL Server</p>

    </main>

</body>
</html>
