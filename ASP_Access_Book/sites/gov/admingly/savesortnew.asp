<!--#include file="../mzfunc.asp"-->
<!--#include file="admintop.asp"-->
<!--#include file="isadmin.asp"-->
<%  dim frominto,title,body
      if request.form("ccc")="�ύ" then
     frominto=request.Form("sort")
	 title=request.form("title")
	 body=server.HTMLEncode(request.form("body"))
	   if frominto="" or title="" or body="" then
	   response.write"<script>alert('����������ϸ');history.back();</script>"
	            else
		sql="select * from allarti where title='"&title&"' and frominto="&frominto
		    set rs=conn.execute(sql)
			if not rs.eof then
		response.write"<script>alert('�����ű����Ѵ���,��������ű���');history.back();</script>"		
	            else
	 sql="insert into allarti(title,body,frominto,times) values('"&title&"','"&body&"','"&frominto&"','"&now()&"')"
         conn.execute(sql)
		 end if
		 end if
		 else
		 response.write"<script>alert('�Ƿ�����');history.back();</script>"
            end if
%>
<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr>
    <td height="30" bgcolor="#B4E7F1"><div align="center"><strong>��ӳɹ�</strong></div></td>
  </tr>
  <tr>
    <td><div align="center">��Ŀ�����Ѿ��ɹ����,����Է�����ҳ���в鿴!</div></td>
  </tr>
  <tr>
    <td><div align="center"><a href="sort_new.asp">�����ء�</a><a href="adminbody.asp">����������������</a></div></td>
  </tr>
</table>
<!--#include file="adminfoot.asp"-->
</body>
</html>
