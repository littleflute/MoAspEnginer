<!--#include file="rscoon_1.asp"-->
<!--#include file="isadmin.asp"-->
<html>
<head>
<META content="text/html; charset=gb2312" http-equiv=Content-Type><LINK 
href="../xsmz.css" rel=stylesheet>
</head>
<body>
<%
     dim allcount,id,i,errorcount
	   errorcount=0
	 allcount=trim(request.querystring("allcount"))
	 id=trim(request.querystring("id"))
         dim votes(9)
	 if allcount="" or id="" then
	 response.write"<script>alert('�Բ���,�������ش���');history.back();</script>"
           errorcount=1
		   end if
     for i=1 to allcount
	 if trim(request.form("vote"&i))="" then
    response.write"<script>alert('���ͶƱ��Ŀ��д����');history.back();</script>"
      errorcount=1
         else
		 votes(i)=trim(request.form("vote"&i))
		  end if
		  next
		  if errorcount=0 then
		  for i=1 to allcount
		  sql="insert into votes(votetitle,fromid) values('"&votes(i)&"',"&id&")"
		     conn.execute(sql)
			 next
%>
<table width="100%" border="0" align="center" bgcolor="#EFEFEF">
  <tr>
    <td bgcolor="#B1E1F3">
<div align="center"><strong>��ӳɹ�</strong></div></td>
  </tr>
  <tr>
    <td>��رմ���,�ڸ����ڵ����ˢ�¡��ٵ��������ӵ���,�鿴����!</td>
  </tr>
  <tr>
    <td><div align="center"><a href="javascript:window.close()">���رմ��ڡ�</a></div></td>
  </tr>
</table>
<% end if %>
</body>
</html>
