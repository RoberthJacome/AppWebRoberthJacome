<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<% Option Explicit %>
<%
Response.CodePage = 65001
Response.Charset = "utf-8"
%>
<!--#include file="includes/conexion.asp"-->
<%
' ---------- Editar canción (solo admin) ----------
RequiereAdmin

Dim conn, cmd, rs, rsArt, idSong, nombre, ruta, artistaSel, errorMsg, nombreUsuario, etiquetaRol
Dim tmpId

nombreUsuario = Session("Username") & ""
etiquetaRol   = "Administrador"

' El id llega por GET (enlace) o por POST (campo oculto)
idSong = 0
tmpId = Request.QueryString("id") & ""
If tmpId <> "" And IsNumeric(tmpId) Then idSong = CLng(tmpId)
tmpId = Request.Form("id") & ""
If tmpId <> "" And IsNumeric(tmpId) Then idSong = CLng(tmpId)
If idSong <= 0 Then Response.Redirect "songs.asp"

errorMsg = ""

' --- Datos actuales de la canción ---
Set conn = AbrirConexion()

Set cmd = Server.CreateObject("ADODB.Command")
Set cmd.ActiveConnection = conn
cmd.CommandType = adCmdText
cmd.CommandText = "SELECT s.SONG_NAME, s.SONG_PATH, " & _
                  "       (SELECT TOP 1 p.ID_ARTIST FROM TBL_PLAYER p WHERE p.ID_SONG = s.ID_SONG) AS ID_ARTIST " & _
                  "FROM TBL_SONG s WHERE s.ID_SONG = ?"
cmd.Parameters.Append cmd.CreateParameter("pId", adInteger, adParamInput, , idSong)

Set rs = cmd.Execute
If rs.EOF Then
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
    Response.Redirect "songs.asp"
End If

nombre     = rs("SONG_NAME") & ""
ruta       = rs("SONG_PATH") & ""
artistaSel = 0
If Not IsNull(rs("ID_ARTIST")) Then artistaSel = CLng(rs("ID_ARTIST"))

rs.Close
Set rs = Nothing

' --- Guardar cambios ---
If UCase(Request.ServerVariables("REQUEST_METHOD")) = "POST" Then

    nombre = Trim(Request.Form("nombre") & "")
    ruta   = Trim(Request.Form("ruta") & "")
    tmpId  = Request.Form("artista") & ""
    If tmpId <> "" And IsNumeric(tmpId) Then artistaSel = CLng(tmpId)

    If nombre = "" Then
        errorMsg = "El nombre de la canción es obligatorio."
    ElseIf Len(ruta) <= 12 Or LCase(Left(ruta, 12)) <> "~/songfiles/" Then
        errorMsg = "La ruta debe empezar con ~/songFiles/ (ejemplo: ~/songFiles/archivo.mp3)."
    ElseIf InStr(ruta, "..") > 0 Then
        errorMsg = "La ruta no puede contener '..'."
    ElseIf artistaSel <= 0 Then
        errorMsg = "Selecciona un artista."
    End If

    If errorMsg = "" Then
        On Error Resume Next
        conn.BeginTrans

        Set cmd = Server.CreateObject("ADODB.Command")
        Set cmd.ActiveConnection = conn
        cmd.CommandType = adCmdText
        cmd.CommandText = "SET NOCOUNT ON; " & _
                          "UPDATE TBL_SONG SET SONG_NAME = ?, SONG_PATH = ? WHERE ID_SONG = ?; " & _
                          "DELETE FROM TBL_PLAYER WHERE ID_SONG = ?; " & _
                          "INSERT INTO TBL_PLAYER (ID_SONG, ID_ARTIST) VALUES (?, ?);"
        cmd.Parameters.Append cmd.CreateParameter("pNombre", adVarWChar, adParamInput, 150, nombre)
        cmd.Parameters.Append cmd.CreateParameter("pRuta", adVarWChar, adParamInput, 300, ruta)
        cmd.Parameters.Append cmd.CreateParameter("pId1", adInteger, adParamInput, , idSong)
        cmd.Parameters.Append cmd.CreateParameter("pId2", adInteger, adParamInput, , idSong)
        cmd.Parameters.Append cmd.CreateParameter("pId3", adInteger, adParamInput, , idSong)
        cmd.Parameters.Append cmd.CreateParameter("pArtista", adInteger, adParamInput, , artistaSel)
        cmd.Execute

        If Err.Number <> 0 Then
            errorMsg = "No se pudieron guardar los cambios."
            Err.Clear
            conn.RollbackTrans
        Else
            conn.CommitTrans
        End If
        On Error GoTo 0

        If errorMsg = "" Then
            conn.Close
            Set conn = Nothing
            Response.Redirect "songs.asp?ok=edit"
        End If
    End If

End If

' --- Artistas para el <select> ---
Set rsArt = conn.Execute("SELECT ID_ARTIST, CONCAT(ARTIST_NAME, ' ', ARTIST_LASTNAME) AS NOMBRE " & _
                         "FROM TBL_ARTIST ORDER BY ARTIST_NAME")
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Editar canción · PoliMusic</title>
<link rel="stylesheet" href="css/estilos.css">
</head>
<body>

<div class="app">

    <aside class="sidebar">
        <a class="brand" href="songs.asp">
            <span class="brand-mark">♪</span>
            <span class="brand-text">Poli<span>Music</span></span>
        </a>

        <nav class="nav">
            <a class="nav-item" href="songs.asp">
                <svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
                    <path d="M9 18V6l10-2v11"/><circle cx="6.5" cy="18" r="2.5"/><circle cx="16.5" cy="15" r="2.5"/>
                </svg>
                Song List
            </a>
            <a class="nav-item" href="song_add.asp">
                <svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
                    <path d="M12 5v14M5 12h14"/>
                </svg>
                Agregar canción
            </a>
        </nav>

        <div class="side-foot">
            <div class="user-card">
                <span class="avatar"><%= Server.HTMLEncode(UCase(Left(nombreUsuario, 1))) %></span>
                <span class="user-meta">
                    <strong><%= Server.HTMLEncode(nombreUsuario) %></strong>
                    <small><%= etiquetaRol %></small>
                </span>
            </div>
            <a class="btn btn-ghost btn-block" href="logout.asp">Cerrar sesión</a>
        </div>
    </aside>

    <main class="main">

        <header class="topbar">
            <div class="welcome">
                <small>Sesión iniciada</small>
                <strong>Bienvenido, <%= Server.HTMLEncode(nombreUsuario) %></strong>
            </div>
            <span class="pill pill-administrador">Administrador</span>
        </header>

        <a class="back-link" href="songs.asp">← Volver a Song List</a>

        <section class="hero">
            <div>
                <h1>Editar canción</h1>
                <p class="count">Modifica los datos de <strong><%= Server.HTMLEncode(nombre) %></strong>.</p>
            </div>
        </section>

        <% If errorMsg <> "" Then %>
        <div class="alert" role="alert">
            <span class="alert-ico">!</span>
            <span><%= Server.HTMLEncode(errorMsg) %></span>
        </div>
        <% End If %>

        <div class="panel form-panel">
            <form method="post" action="song_edit.asp" autocomplete="off">

                <input type="hidden" name="id" value="<%= idSong %>">

                <div class="field">
                    <label for="nombre">Nombre de la canción</label>
                    <input class="input" type="text" id="nombre" name="nombre" maxlength="150"
                           value="<%= Server.HTMLEncode(nombre) %>"
                           placeholder="Ej. Adventure" required autofocus>
                </div>

                <div class="field">
                    <label for="ruta">Ruta del archivo</label>
                    <input class="input" type="text" id="ruta" name="ruta" maxlength="300"
                           value="<%= Server.HTMLEncode(ruta) %>"
                           placeholder="~/songFiles/archivo.mp3" required>
                    <p class="hint">Formato obligatorio: <code>~/songFiles/archivo.mp3</code></p>
                </div>

                <div class="field">
                    <label for="artista">Artista</label>
                    <select class="input" id="artista" name="artista" required>
                        <option value="">Selecciona un artista…</option>
                        <%
                        Do While Not rsArt.EOF
                        %>
                        <option value="<%= rsArt("ID_ARTIST") %>"<% If artistaSel = CLng(rsArt("ID_ARTIST")) Then Response.Write " selected" %>><%= Server.HTMLEncode(rsArt("NOMBRE") & "") %></option>
                        <%
                            rsArt.MoveNext
                        Loop
                        %>
                    </select>
                </div>

                <div class="form-actions">
                    <button type="submit" class="btn btn-primary">Guardar cambios</button>
                    <a class="btn btn-ghost" href="songs.asp">Cancelar</a>
                </div>

            </form>
        </div>

    </main>
</div>

<%
rsArt.Close
Set rsArt = Nothing
conn.Close
Set conn = Nothing
%>

</body>
</html>
