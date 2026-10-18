<%
' ============================================================
'  PoliMusic - Conexión a SQL Server
'  ÚNICO lugar donde vive la cadena de conexión.
' ============================================================

' Contenedor SQL Server 2025 Express en localhost,14333.
' El contenedor usa certificado autofirmado -> TrustServerCertificate=yes
Const CADENA_CONEXION = "Provider=MSOLEDBSQL19;Server=localhost,14333;Database=BDD_PoliMusic;UID=sa;PWD=Politecnica1;TrustServerCertificate=yes;"

' Plan B: algunas versiones de MSOLEDBSQL19 ignoran TrustServerCertificate contra
' certificados autofirmados. En ese caso se desactiva el cifrado (solo red local).
Const CADENA_CONEXION_ALT1 = "Provider=MSOLEDBSQL19;Server=localhost,14333;Database=BDD_PoliMusic;UID=sa;PWD=Politecnica1;Use Encryption for Data=False;"

' Plan C: driver 18.x (Microsoft OLE DB Driver for SQL Server).
Const CADENA_CONEXION_ALT2 = "Provider=MSOLEDBSQL;Server=localhost,14333;Database=BDD_PoliMusic;UID=sa;PWD=Politecnica1;TrustServerCertificate=yes;"

' NO uses SQLOLEDB: es obsoleto y falla contra SQL Server 2025.

' Constantes ADO (así no hace falta adovbs.inc)
Const adCmdText       = 1
Const adParamInput    = 1
Const adVarWChar      = 202
Const adInteger       = 3
Const adUseClient     = 3
Const adOpenStatic    = 3
Const adLockReadOnly  = 1

' Abre la conexión probando la cadena principal y, si falla, los planes B y C.
Function AbrirConexion()
    Dim cn, cadenas(2), i, ultimoError

    cadenas(0) = CADENA_CONEXION
    cadenas(1) = CADENA_CONEXION_ALT1
    cadenas(2) = CADENA_CONEXION_ALT2

    ultimoError = ""

    For i = 0 To 2
        Set cn = Server.CreateObject("ADODB.Connection")
        cn.ConnectionTimeout = 15

        On Error Resume Next
        cn.Open cadenas(i)
        If Err.Number = 0 Then
            On Error GoTo 0
            Set AbrirConexion = cn
            Exit Function
        End If
        ultimoError = Err.Description
        Err.Clear
        On Error GoTo 0

        Set cn = Nothing
    Next

    Response.Write "<div style=""max-width:640px;margin:56px auto;padding:26px 30px;border-radius:16px;background:#2A2A2F;color:#F3F3F4;font-family:Segoe UI,Arial,sans-serif;border-left:4px solid #D00000;"">"
    Response.Write "<strong>No se pudo conectar con la base de datos BDD_PoliMusic.</strong><br><br>"
    Response.Write "Comprueba que el contenedor de SQL Server esté activo en <b>localhost,14333</b> y que el driver OLE DB esté instalado."
    Response.Write "<br><br><span style=""font-size:12px;opacity:.65"">" & Server.HTMLEncode(ultimoError) & "</span></div>"
    Response.End
End Function

' "~/songFiles/x.mp3" -> "songFiles/x.mp3" (el "~/" no existe en ASP clásico)
Function RutaAudio(valor)
    If IsNull(valor) Then
        RutaAudio = ""
    Else
        RutaAudio = Replace(CStr(valor), "~/", "")
    End If
End Function

' Exige sesión iniciada
Sub RequiereSesion()
    If Trim(Session("Username") & "") = "" Then
        Response.Redirect "login.asp"
    End If
End Sub

' Exige sesión de administrador (USER_TYPE = 1). Se valida en el servidor.
Sub RequiereAdmin()
    RequiereSesion

    If Not IsNumeric(Session("UserType")) Then
        Response.Redirect "songs.asp"
    End If

    If CLng(Session("UserType")) <> 1 Then
        Response.Redirect "songs.asp"
    End If
End Sub
%>
