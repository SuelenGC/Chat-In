<% 

Function FormataHora(DataHora)
  FormataHora = Right("0" & DatePart("d",DataHora),2) & "/" & Right("0" & DatePart("m",DataHora),2) & "/" & DatePart("yyyy",DataHora) & " " & Right("0" & DatePart("H",DataHora),2) & ":" & Right("0" & DatePart("n",DataHora),2) & ":" & Right("0" & DatePart("s",DataHora),2)
End Function


IF Request.Cookies("Apelido")="" and request.form("opiniao")<>"" THEN 
    if request.form("apelido")<>"" then 
        Response.Cookies("Apelido")=Request.Form("Apelido") 
        EnviaMensagem()
    else 
        'Response.Cookies("Apelido")="Anonimo" 
        alertas = "Preencha o apelido."
    end if 

END IF 

function EnviaMensagem()
    Application.Lock 
    Temporal=Application("Opinioes") 
    
    FOR i=7 TO 0 STEP -1 
        Temporal(i+1)=Temporal(i) 
    NEXT 
    
    if request.form("apelido")<>"" then 
        Temporal(0)="<FONT COLOR=""#000000"">" & Request.Form("Apelido") & " entrou no minichat</FONT>" 
    else 
        Temporal(0)="<FONT COLOR=""#000000"">Anonimo entro no minichat</FONT>" 
    end if 
    
    Application("Opinioes")=Temporal 
    Application.Unlock 
end function

IF Request.Form("Opiniao")<>"" THEN 
    Apelido=Request.Cookies("Apelido") 
    Application.Lock 

    Temporal=Application("Opinioes") 
        
    FOR i=7 TO 0 STEP -1 
        Temporal(i+1)=Temporal(i) 
    NEXT 
    
    Temporal(0)= "<b>" & FormataHora(now()) & " " & Apelido & " diz - </b>" & Request.Form("Opiniao") 

    Application("Opinioes") = Temporal 

    FOR Opinion=8 to 0 step -1 
        mensagens = mensagens & "\n" & Opinion & " - " & Application("Opinioes")(Opinion)
    NEXT 

    'response.write("<script language='javascript'>alert('" & mensagens & "')</script>")

    Application.Unlock 
END IF
%> 

<html> 

<head> 
<title>incluir opiniao</title> 
<base target="_self"> 
</head> 

<body bgcolor="#6699FF"> 
    <FORM METHOD="POST" ACTION="incluir.asp" color="#FFFFFF"> 
    
        <%= alertas %>
        <% IF Request.Cookies("Apelido")="" THEN %> 
            Apelido:<INPUT TYPE="TEXT" SIZE=10 NAME="Apelido"> 
            <input type="hidden" name="go" size="20" value="si"><BR> 
        <% END IF %> 
        Msg: <INPUT TYPE="TEXT" SIZE=30 NAME="Opiniao"> 
        <INPUT TYPE="SUBMIT" VALUE="Enviar">         
    
        <a href="fechar.asp" target="_top">Sair</a> 
        <% if Request.Cookies("Apelido") = "Trillian" then %>
            <a href="LimparApp.asp" target="_top">Limpar </a> 
        <% end if %>
    </FORM> 
</body> 

</html> 