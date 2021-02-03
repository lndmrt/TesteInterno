	<%@ LANGUAGE=VBScript %>

	<% Option explicit%>

<!--#include file="bib_conexao.asp"-->

<%

dim rs, conexao, strsql, var_codigo

call abre_conexao

strsql ="select distinct cli_nome, con_codigo, con_id, con_modelo from view_aah"
set rs=conexao.execute(strsql)
if not rs.eof then
var_codigo=rs("con_codigo")
end if
%>
<html>
<head>
<title>expedicao</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<body bgcolor="#CCCCCC">
<table width="750" border="1" align="center" bgcolor="#FFFFFF">
  <tr align="left" valign="top"> 
    <td height="47" align="center"><a href="menu.asp"><img src="Images/cabeca.gif" width="750" height="50" border="0"></a></td>
  </tr>
  <tr align="left" valign="top"> 
    <td height="19" bgcolor="#FF0000"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#FFFFFF">&gt; 
      FICHA DE LIBERA&Ccedil;&Atilde;O DE EQUIPAMENTOS - CLIQUE NO LINK CORRESPONDENTE 
      PARA DETALHES</font></b></td>
  </tr>
  <tr align="left" valign="top"> 
    <td height="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#FF9900"><b>Cliente</b></font><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#FF9900"></font></td>
  </tr>
  <% if not rs.eof then
  while not rs.eof %>
  <tr align="left" valign="top"> 
    <td height="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#FF9900"><a href="expedicao_cont_aah.asp?codigo=<%=rs("con_codigo")%>&id=<%=rs("con_id")%>&modelo=<%=rs("con_modelo")%>"><%=rs("cli_nome")%></a> 
      </font></td>
  </tr>
  <% rs.movenext
  wend
  end if %>
</table>
<br>

</body>
</html>
caiqueLopescriado
<% set rs=nothing
call fecha_conexao %>
