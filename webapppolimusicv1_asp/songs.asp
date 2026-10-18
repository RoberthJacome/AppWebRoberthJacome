<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<% Option Explicit %>
<%
Response.CodePage = 65001
Response.Charset = "utf-8"
%>
<!--#include file="includes/conexion.asp"-->
<%
' ---------- Lista de canciones ----------
RequiereSesion

Dim conn, rs, sql, esAdmin, nombreUsuario, etiquetaRol, colSpan
Dim i, total, ruta, artistas, albumes, plays, idSong, cancion, flash, ok
Dim flashClass, flashIco

esAdmin       = (CLng(Session("UserType")) = 1)
nombreUsuario = Session("Username") & ""
etiquetaRol   = "Oyente"
If esAdmin Then etiquetaRol = "Administrador"

' Aviso de resultado (viene de add / edit / delete)
flash      = ""
flashClass = "alert alert-ok"
flashIco   = "✓"
ok = LCase(Request.QueryString("ok") & "")
If ok = "add"  Then flash = "Canción agregada correctamente."
If ok = "edit" Then flash = "Canción actualizada correctamente."
If ok = "del"  Then flash = "Canción eliminada del catálogo."
If ok = "err" Then
    flash      = "No se pudo eliminar la canción."
    flashClass = "alert"
    flashIco   = "!"
End If

' Una sola consulta: canciones + artistas + álbumes agrupados con STRING_AGG
sql = "SELECT s.ID_SONG, s.SONG_NAME, s.SONG_PATH, s.PLAYS, " & _
      "       ISNULL(ar.ARTISTAS, '') AS ARTISTAS, " & _
      "       ISNULL(al.ALBUMES, '')  AS ALBUMES " & _
      "FROM TBL_SONG s " & _
      "LEFT JOIN ( " & _
      "    SELECT p.ID_SONG, STRING_AGG(CONCAT(a.ARTIST_NAME, ' ', a.ARTIST_LASTNAME), ', ') AS ARTISTAS " & _
      "    FROM TBL_PLAYER p " & _
      "    INNER JOIN TBL_ARTIST a ON a.ID_ARTIST = p.ID_ARTIST " & _
      "    GROUP BY p.ID_SONG " & _
      ") ar ON ar.ID_SONG = s.ID_SONG " & _
      "LEFT JOIN ( " & _
      "    SELECT t.ID_SONG, STRING_AGG(b.TITLE_ALBUM, ', ') AS ALBUMES " & _
      "    FROM TBL_TRACKLIST t " & _
      "    INNER JOIN TBL_ALBUM b ON b.ID_ALBUM = t.ID_ALBUM " & _
      "    GROUP BY t.ID_SONG " & _
      ") al ON al.ID_SONG = s.ID_SONG " & _
      "ORDER BY s.SONG_NAME"

Set conn = AbrirConexion()

' Cursor estático en el cliente para conocer el total antes de pintar
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = adUseClient
rs.Open sql, conn, adOpenStatic, adLockReadOnly
total = rs.RecordCount

colSpan = 6
If esAdmin Then colSpan = 7
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Song List · PoliMusic</title>
<link rel="stylesheet" href="css/estilos.css">
</head>
<body>

<div class="app">

    <!-- ================= Sidebar ================= -->
    <aside class="sidebar">
        <a class="brand" href="songs.asp">
            <span class="brand-mark">♪</span>
            <span class="brand-text">Poli<span>Music</span></span>
        </a>

        <nav class="nav">
            <a class="nav-item active" href="songs.asp">
                <svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
                    <path d="M9 18V6l10-2v11"/><circle cx="6.5" cy="18" r="2.5"/><circle cx="16.5" cy="15" r="2.5"/>
                </svg>
                Song List
            </a>

            <% If esAdmin Then %>
            <a class="nav-item" href="song_add.asp">
                <svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
                    <path d="M12 5v14M5 12h14"/>
                </svg>
                Agregar canción
            </a>
            <% End If %>
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

    <!-- ================= Contenido ================= -->
    <main class="main">

        <header class="topbar">
            <div class="welcome">
                <small>Sesión iniciada</small>
                <strong>Bienvenido, <%= Server.HTMLEncode(nombreUsuario) %></strong>
            </div>
            <span class="pill<%= " " & LCase(Replace(etiquetaRol, " ", "-")) %>"><%= etiquetaRol %></span>
        </header>

        <section class="hero">
            <div>
                <h1>Song List</h1>
                <p class="count">
                    <%= total %> canciones en el catálogo
                </p>
            </div>
            <% If esAdmin Then %>
            <a class="btn btn-primary" href="song_add.asp">
                <span class="plus">+</span> Agregar canción
            </a>
            <% End If %>
        </section>

        <% If flash <> "" Then %>
        <div class="<%= flashClass %>" role="status">
            <span class="alert-ico"><%= flashIco %></span>
            <span><%= Server.HTMLEncode(flash) %></span>
        </div>
        <% End If %>

        <div class="panel">
            <table class="song-table" id="songTable">
                <thead>
                    <tr>
                        <th class="col-num">#</th>
                        <th>Canción</th>
                        <th class="col-artist">Artista(s)</th>
                        <th class="col-album">Álbum(es)</th>
                        <th class="col-plays">Reproducciones</th>
                        <th class="col-play"><span class="sr-only">Reproducir</span></th>
                        <% If esAdmin Then %><th class="col-actions">Acciones</th><% End If %>
                    </tr>
                </thead>
                <tbody>
                <%
                i = 0
                Do While Not rs.EOF
                    i = i + 1
                    idSong   = rs("ID_SONG")
                    cancion  = rs("SONG_NAME") & ""
                    ruta     = RutaAudio(rs("SONG_PATH"))
                    artistas = rs("ARTISTAS") & ""
                    albumes  = rs("ALBUMES") & ""
                    plays    = 0
                    If Not IsNull(rs("PLAYS")) Then plays = CLng(rs("PLAYS"))
                    If artistas = "" Then artistas = "Sin artista"
                    If albumes  = "" Then albumes  = "Sin álbum"
                %>
                    <tr data-src="<%= Server.HTMLEncode(ruta) %>"
                        data-title="<%= Server.HTMLEncode(cancion) %>"
                        data-artist="<%= Server.HTMLEncode(artistas) %>">
                        <td class="col-num"><%= i %></td>
                        <td>
                            <div class="song-cell">
                                <span class="cover">♪</span>
                                <span class="song-meta">
                                    <span class="song-name"><%= Server.HTMLEncode(cancion) %></span>
                                    <span class="song-sub"><%= Server.HTMLEncode(artistas) %></span>
                                </span>
                            </div>
                        </td>
                        <td class="col-artist"><%= Server.HTMLEncode(artistas) %></td>
                        <td class="col-album"><%= Server.HTMLEncode(albumes) %></td>
                        <td class="col-plays"><%= plays %></td>
                        <td class="col-play">
                            <button type="button" class="btn-play" aria-label="Reproducir <%= Server.HTMLEncode(cancion) %>">
                                <svg class="ico-play" viewBox="0 0 24 24" aria-hidden="true"><path d="M8 5v14l11-7z"/></svg>
                                <svg class="ico-pause" viewBox="0 0 24 24" aria-hidden="true"><path d="M6 5h4v14H6zM14 5h4v14h-4z"/></svg>
                            </button>
                        </td>
                        <% If esAdmin Then %>
                        <td class="col-actions">
                            <div class="actions">
                                <a class="link-edit" href="song_edit.asp?id=<%= idSong %>">Editar</a>
                                <a class="link-del" href="song_delete.asp?id=<%= idSong %>"
                                   onclick="return confirm('¿Eliminar esta canción del catálogo?');">Eliminar</a>
                            </div>
                        </td>
                        <% End If %>
                    </tr>
                <%
                    rs.MoveNext
                Loop

                rs.Close
                Set rs = Nothing
                conn.Close
                Set conn = Nothing
                %>
                <% If total = 0 Then %>
                    <tr>
                        <td colspan="<%= colSpan %>" class="empty">
                            Todavía no hay canciones en el catálogo.
                        </td>
                    </tr>
                <% End If %>
                </tbody>
            </table>
        </div>

    </main>
</div>

<!-- ================= Reproductor fijo ================= -->
<div class="player" id="player">

    <div class="player-track">
        <div class="player-cover">♪</div>
        <div class="player-meta">
            <span class="player-title" id="playerTitle">Selecciona una canción</span>
            <span class="player-artist" id="playerArtist">PoliMusic</span>
        </div>
    </div>

    <div class="player-center">
        <div class="player-controls">
            <button type="button" class="ctrl" id="btnPrev" aria-label="Anterior">
                <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M6 6h2v12H6zm3.5 6L18 18V6z"/></svg>
            </button>
            <button type="button" class="ctrl ctrl-main" id="btnToggle" aria-label="Reproducir">
                <svg class="ico-play" viewBox="0 0 24 24" aria-hidden="true"><path d="M8 5v14l11-7z"/></svg>
                <svg class="ico-pause" viewBox="0 0 24 24" aria-hidden="true"><path d="M6 5h4v14H6zM14 5h4v14h-4z"/></svg>
            </button>
            <button type="button" class="ctrl" id="btnNext" aria-label="Siguiente">
                <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M16 6h2v12h-2zM6 6l8.5 6L6 18z"/></svg>
            </button>
        </div>
        <div class="player-progress">
            <span class="time" id="timeCurrent">0:00</span>
            <input type="range" id="seek" min="0" max="100" step="0.1" value="0" aria-label="Progreso de la canción">
            <span class="time" id="timeTotal">0:00</span>
        </div>
    </div>

    <div class="player-right">
        <span class="pill pill-soft">ASP · SQL Server</span>
    </div>

</div>

<audio id="audio" preload="metadata"></audio>

<script src="js/player.js"></script>
</body>
</html>
