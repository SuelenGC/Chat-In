<% PaginaAtual="http://"&_ 
Request.ServerVariables("SERVER_NAME")&_ 
Request.ServerVariables("SCRIPT_NAME") %> 
<html> 

<head> 
<META HTTP-EQUIV="REFRESH" CONTENT="3;<%=PaginaActual%>"> 
<title>MiniChat (visualizacao)</title> 
</head> 

<body> 
<FONT FACE="Calibri" COLOR="black" size="1"> 

<% 
IF NOT isArray( Application("Opinioes")) THEN 
    Application.Lock 
    Dim Auxiliar() 
    Redim Auxiliar(9) 
    Application("Opinioes")=Auxiliar 
    Application.UnLock
END IF 

Temporal=Application("Opinioes") 

FOR Opinion=8 to 0 step -1 
%> 

<%= Temporal(Opinion)%> <BR> 

<% NEXT %> 
<FONT> 
</body> 
</html> 