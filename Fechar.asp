<% if Request.cookies("Apelido")<>"" then 
Application.Lock 
Temporal=Application("Opinioes") 
FOR i=7 TO 0 STEP -1 
Temporal(i+1)=Temporal(i) 
NEXT 
Temporal(0)="<FONT COLOR=""#FF0000"">" &Request.cookies("Apelido")&" saiu do minichat</FONT>" 
Application("Opinioes")=Temporal 
Application.Unlock 
response.cookies("apelido")="" 

END IF%> 
<HTML> 
<HEAD> 
<script language="JavaScript">
    { close(); } 
</SCRIPT> 
</HEAD> 
<BODY> 
</BODY> 
</HTML> 