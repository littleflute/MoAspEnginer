<html>
<head>
<LINK href="style.css" rel=stylesheet>
<base target="right">
</head>
<body bgcolor="#009458">
<%
if session("admin")="" and session("flag")="" then
	response.write "Error!!"
else
	call main()
end if
sub main()
%>
<br>
<br>
<table border=0 width=98% cellspacing="1" cellpadding="1" align="center">

  <tr> 
    <td width=100 align=center bgcolor="#cce6ed"><FONT style=line-height:150%> 
      ��<A href="freeadd.asp" target="right" title="ֱ��������ص�ַ����Ϣ">���ӰƬ</A></FONT></td>
  </tr>
  <tr> 
    <td width=100 align=center bgcolor="#a5d0dc"><FONT style=line-height:150%> 
      <%if session("flag")<3 then%>
      ��<A href="adminedit.asp" target="right">�޸�ɾ��</A></FONT></td>
  </tr>
  <tr> 
    <td width=100 align=center bgcolor="#cce6ed"><FONT style=line-height:150%> 
      <%if session("flag")=1 then%>
      ��<A href="classmana.asp" target="right">��Ŀ����</A></FONT></td>
  </tr>
  <tr> 
    <td width=100 align=center bgcolor="#a5d0dc"><FONT style=line-height:150%>��<A href="adminuser.asp" target="right">�û�����</A></FONT></td>
  </tr>
  <tr> 
    <td width=100 align=center bgcolor="#a5d0dc"><font style=line-height:150%>��</font><a href="adminuser1.asp" target="right">��Ա����</a></td>
  </tr>

  <%end if%>
  <%end if%>
  <tr> 
    <td width=100 align=center bgcolor="#a5d0dc" height="21"><FONT style=line-height:150%>��<a href="main.asp">������ҳ</a></FONT></td>
  </tr>
  <tr> 
    <td width=100 align=center bgcolor="#cce6ed"><FONT style=line-height:150%>��<A href="logout.asp" target=_top>�˳�ϵͳ</A></FONT></td>
  </tr>
</table>
<%
end sub
%>
</html>