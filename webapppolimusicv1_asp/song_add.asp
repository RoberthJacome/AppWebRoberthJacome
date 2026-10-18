<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<% Option Explicit %>
<%
Response.CodePage = 65001
Response.Charset = "utf-8"
%>
<!--#include file="includes/conexion.asp"-->
<%
' ---------- Agregar canción (solo admin) ----------
RequiereAdmin

Dim conn, cmd, rs, rsArt, nombre, ruta, idArtista, errorMsg, nuevoId, nombreUsuario, etiquetaRol
Dim tmpArt

nombreUsuario = Session("Username") & ""
etiquetaRol   = "Administrador"

nombre    = ""
ruta      = "~/songFiles/"
idArtista = 0
errorMsg  = ""
nuevoId   = 0

If UCase(Request.ServerVariables("REQUEST_METHOD")) = "POST" Then

    nombre = Trim(Request.Form("nombre") & "")
    ruta   = Trim(Request.Form("ruta") & "")
    tmpArt = Request.Form("artista") & ""
    If tmpArt <> "" And IsNumeric(tmpArt) Then idArtista = CLng(tmpArt)

    ' Validaciones básicas
    If nombre = "" Then
        errorMsg = "El nombre de la canción es obligatorio."
    ElseIf Len(ruta) <= 12 Or LCase(Left(ruta, 12)) <> "~/songfiles/" Then
        errorMsg = "La ruta debe empezar con ~/songFiles/ (ejemplo: ~/songFiles/archivo.mp3)."
    ElseIf InStr(ruta, "..") > 0 Then
        errorMsg = "La ruta no puede contener '..'."
    ElseIf idArtista <= 0 Then
        errorMsg = "Selecciona un artista."
    End If

    If errorMsg = "" Then
        Set conn = AbrirConexion()

        On Error Resume Next
        conn.BeginTrans

        ' 1) Insertar la canción y recuperar su ID
        Set cmd = Server.CreateObject("ADODB.Command")
        Set cmd.ActiveConnection = conn
        cmd.CommandType = adCmdText
        cmd.CommandText = "SET NOCOUNT ON; " & _
                          "INSERT INTO TBL_SONG (SONG_NAME, SONG_PATH, PLAYS) VALUES (?, ?, 0); " & _
                          "SELECT CAST(SCOPE_IDENTITY() AS INT) AS NUEVO_ID;"
        cmd.Parameters.Append cmd.CreateParameter("pNombre", adVarWChar, adParamInput, 150, nombre)
        cmd.Parameters.Append cmd.CreateParameter("pRuta", adVarWChar, adParamInput, 300, ruta)

        Set rs = cmd.Execute
        If Err.Number <> 0 Then
            errorMsg = "No se pudo guardar la canción."
            Err.Clear
        Else
            If Not rs.EOF Then nuevoId = CLng(rs("NUEVO_ID"))
            rs.Close
            Set rs = Nothing
        End If

        ' 2) Relación canción - artista
        If errorMsg = "" And nuevoId > 0 Then
            Set cmd = Server.CreateObject("ADODB.Command")
            Set cmd.ActiveConnection = conn
            cmd.CommandType = adCmdText
            cmd.CommandText = "INSERT INTO TBL_PLAYER (ID_SONG, ID_ARTIST) VALUES (?, ?)"
            cmd.Parameters.Append cmd.CreateParameter("pSong", adInteger, adParamInput, , nuevoId)
            cmd.Parameters.Append cmd.CreateParameter("pArtist", adInteger, adParamInput, , idArtista)
            cmd.Execute

            If Err.Number <> 0 Then
                errorMsg = "No se pudo asociar el artista a la canción."
                Err.Clear
            End If
        ElseIf nuevoId = 0 And errorMsg = "" Then
            errorMsg = "No se pudo obtener el ID de la canción insertada."
        End If

        If errorMsg = "" Then
            conn.CommitTrans
        Else
            conn.RollbackTrans
        End If
        On Error GoTo 0

        conn.Close
        Set conn = Nothing

        If errorMsg = "" Then
            Response.Redirect "songs.asp?ok=add"
        End If
    End If

End If

' Artistas para el <select>
Set conn = AbrirConexion()
Set rsArt = conn.Execute("SELECT ID_ARTIST, CONCAT(ARTIST_NAME, ' ', ARTIST_LASTNAME) AS NOMBRE " & _
                         "FROM TBL_ARTIST ORDER BY ARTIST_NAME")
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Agregar canción · PoliMusic</title>
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
            <a class="nav-item active" href="song_add.asp">
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
                <h1>Agregar canción</h1>
                <p class="count">Registra una canción y su artista en el catálogo.</p>
            </div>
        </section>

        <% If errorMsg <> "" Then %>
        <div class="alert" role="alert">
            <span class="alert-ico">!</span>
            <span><%= Server.HTMLEncode(errorMsg) %></span>
        </div>
        <% End If %>

        <div class="panel form-panel">
            <form method="post" action="song_add.asp" autocomplete="off">

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
                        <option value="<%= rsArt("ID_ARTIST") %>"<% If idArtista = CLng(rsArt("ID_ARTIST")) Then Response.Write " selected" %>><%= Server.HTMLEncode(rsArt("NOMBRE") & "") %></option>
                        <%
                            rsArt.MoveNext
                        Loop
                        %>
                    </select>
                </div>

                <div class="form-actions">
                    <button type="submit" class="btn btn-primary">Guardar canción</button>
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
