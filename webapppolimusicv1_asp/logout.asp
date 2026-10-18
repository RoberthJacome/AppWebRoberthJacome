<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<%
Response.CodePage = 65001
Response.Charset = "utf-8"

' Cerrar sesión y volver al login
Session.Abandon
Response.Redirect "login.asp"
%>
