<!--#include file="rscoon_1.asp"-->
<!--#include file="isadmin.asp"-->
<html>
<head>
<META content="text/html; charset=gb2312" http-equiv=Content-Type><LINK 
href="../xsmz.css" rel=stylesheet>
</head>
<% dim id,errorcount
     errorcount=""
id=request.querystring("id")
  sql="select isvote from stat where id="&id
Set rs = Server.CreateObject("ADODB.RecordSet")
         rs.open sql,conn,1,3
		 if rs("isvote")="true" then
		 rs("isvote")="false"
		 else
		 sql1="select isvote,title from stat where isvote='true'"
		 set rs1=conn.execute(sql1)
		 if not rs1.eof then
		 response.write"<script>alert('�Բ���,�Ѿ������˱��ͶƱ��\nҪ�뿪����ͶƱͳ��,���ȹر��ѿ������Ǹ�ͶƱѡ��');</script>"
		    errorcount="�Ѿ��С�"&rs1("title")&"��ͶƱͳ�ƿ�����,���ȹرմ�ѡ��,���ܴ�������ͶƱѡ��"
		    else
		 rs("isvote")="true"
		 end if 
		 end if
		 rs.update
 if errorcount="" then %>
<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr> 
    <td height="30" bgcolor="#B1E1F3"> <div align="center"><strong><% if rs("isvote")="true" then response.write"����" else response.write"�ر�" end if%>�ɹ�</strong></div></td>
  </tr>
  <tr> 
    <td><div align="center">��رմ���,Ȼ���������ڵġ�ˢ�¡����в鿴.</div></td>
  </tr>
  <tr> 
    <td height="20"> <div align="center"><a href="javascript:window.close()">���رմ��ڡ�</a></div></td>
  </tr>
</table>
<% else %>
<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr> 
    <td height="30" bgcolor="#B1E1F3"> <div align="center"><strong><% if rs("isvote")="true" then response.write"�ر�" else response.write"����" end if%>ʧ��</strong></div></td>
  </tr>
  <tr> 
    <td><div align="center"><%=errorcount%></div></td>
  </tr>
  <tr> 
    <td height="20"> <div align="center"><a href="javascript:window.close()">���رմ��ڡ�</a></div></td>
  </tr>
</table>
<% end if %>
</body>
</html>
