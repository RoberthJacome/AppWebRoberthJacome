<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<% Option Explicit %>
<%
Response.CodePage = 65001
Response.Charset = "utf-8"
%>
<!--#include file="includes/conexion.asp"-->
<%
' ---------- Eliminar canción (solo admin) ----------
RequiereAdmin

Dim conn, cmd, idSong, errorMsg, tmpId

idSong = 0
tmpId = Request.QueryString("id") & ""
If tmpId <> "" And IsNumeric(tmpId) Then idSong = CLng(tmpId)
If idSong <= 0 Then Response.Redirect "songs.asp"

errorMsg = ""

Set conn = AbrirConexion()

On Error Resume Next
conn.BeginTrans

' Primero las tablas relacionadas (llaves foráneas), luego la canción
Set cmd = Server.CreateObject("ADODB.Command")
Set cmd.ActiveConnection = conn
cmd.CommandType = adCmdText
cmd.CommandText = "SET NOCOUNT ON; " & _
                  "DELETE FROM TBL_TRACKLIST     WHERE ID_SONG = ?; " & _
                  "DELETE FROM TBL_PLAYER        WHERE ID_SONG = ?; " & _
                  "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_SONG = ?; " & _
                  "DELETE FROM TBL_SONG          WHERE ID_SONG = ?;"
cmd.Parameters.Append cmd.CreateParameter("pId1", adInteger, adParamInput, , idSong)
cmd.Parameters.Append cmd.CreateParameter("pId2", adInteger, adParamInput, , idSong)
cmd.Parameters.Append cmd.CreateParameter("pId3", adInteger, adParamInput, , idSong)
cmd.Parameters.Append cmd.CreateParameter("pId4", adInteger, adParamInput, , idSong)
cmd.Execute

If Err.Number <> 0 Then
    errorMsg = "No se pudo eliminar la canción."
    Err.Clear
    conn.RollbackTrans
Else
    conn.CommitTrans
End If
On Error GoTo 0

conn.Close
Set conn = Nothing

If errorMsg = "" Then
    Response.Redirect "songs.asp?ok=del"
Else
    Response.Redirect "songs.asp?ok=err"
End If
%>
